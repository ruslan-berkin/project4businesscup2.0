"""Calculate playtest metrics from PostgreSQL, SQLite or archived Render logs."""
import argparse
import json
import os
import sqlite3
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path


def read_records(path):
    path = Path(path)
    if path.suffix in ('.sqlite3', '.db'):
        with sqlite3.connect('file:' + str(path.resolve()) + '?mode=ro', uri=True) as db:
            yield from (json.loads(row[0]) for row in db.execute('SELECT payload FROM events ORDER BY ts'))
        return
    for line in path.read_text().splitlines():
        for prefix in ('TE_EVENT ', 'TE_AI_USAGE '):
            if prefix in line:
                line = line.split(prefix, 1)[1]
                break
        try:
            record = json.loads(line)
            if isinstance(record, dict): yield record
        except (ValueError, TypeError):
            continue


def read_postgres(database_url):
    import psycopg

    with psycopg.connect(database_url, connect_timeout=3) as db:
        db.execute("SET TRANSACTION READ ONLY")
        db.execute("SET LOCAL statement_timeout = '15s'")
        for (payload,) in db.execute("SELECT payload FROM research_events ORDER BY ts, dedupe"):
            yield json.loads(payload)


def summarize(records, as_of):
    records = list(records)
    seen, events, usage = set(), [], []
    for row in records:
        if not isinstance(row.get('ts'), (int, float)) or row['ts'] > as_of: continue
        if row.get('provider') and 'usage_available' in row:
            usage.append(row); continue
        if not row.get('event_id') or not row.get('visitor_id'): continue
        key = ((row['visitor_id'], row.get('adventure_id'), row['name'], row.get('turn'))
               if row.get('name') in ('adventure_started', 'turn_completed', 'adventure_completed') else row['event_id'])
        if key in seen: continue
        seen.add(key); events.append(row)
    excluded = {e['visitor_id'] for e in events if e.get('cohort') == 'demo' or e.get('source') in ('team', 'jury')}
    enrolled = {e['visitor_id'] for e in events if e.get('name') == 'enrolled' and e.get('cohort') == 'pilot'
                and e.get('eligible') is True} - excluded
    pilot = [e for e in events if e['visitor_id'] in enrolled and e.get('eligible') is True and e.get('cohort') == 'pilot']
    users = lambda name: {e['visitor_id'] for e in pilot if e['name'] == name}
    turns = defaultdict(list)
    for e in pilot:
        if e['name'] == 'turn_completed': turns[e['visitor_id']].append(e['ts'])
    def retention(start_hour, end_hour):
        eligible = {v: min(times) for v, times in turns.items() if min(times) + end_hour * 3600 <= as_of}
        returned = sum(any(first + start_hour * 3600 <= t < first + end_hour * 3600 for t in turns[v])
                       for v, first in eligible.items())
        return {'eligible_browsers': len(eligible), 'returned_browsers': returned,
                'rate': round(returned / len(eligible), 4) if eligible else None,
                'pending_browsers': len(turns) - len(eligible)}
    enrolled_at = {v:min(e['ts'] for e in pilot if e['name']=='enrolled' and e['visitor_id']==v) for v in enrolled}
    completion_mature = {v:t for v,t in enrolled_at.items() if t+7*86400 <= as_of}
    completion_returned = {e['visitor_id'] for e in pilot if e['name']=='adventure_completed'
                           and e.get('mode')=='free' and e['visitor_id'] in completion_mature
                           and completion_mature[e['visitor_id']] <= e['ts'] < completion_mature[e['visitor_id']]+7*86400}
    interest = {}
    feedback = {}
    for e in sorted(pilot, key=lambda e:e['ts']):
        if e['name'] == 'premium_interest': interest[e['visitor_id']] = e['interest']
        if e['name'] == 'feedback': feedback[e['visitor_id']] = {'rating':e['rating'],'reason':e['reason']}
    usage_summary = {}
    for u in usage:
        key = u['provider'] + '/' + u.get('model', 'unknown')
        totals = usage_summary.setdefault(key, {'responses':0,'usage_missing':0,'input_tokens':0,'output_tokens':0})
        totals['responses'] += 1
        if not u.get('usage_available'): totals['usage_missing'] += 1
        else:
            totals['input_tokens'] += u['input_tokens']; totals['output_tokens'] += u['output_tokens']
    return {'as_of_utc':datetime.fromtimestamp(as_of,timezone.utc).isoformat(),
            'notes':['Self-selected browser IDs, not verified people. Consent and complete event coverage required.',
                     'Jury/team/demo IDs excluded. D7: 168–192h; D30: 720–744h after first observed successful turn.',
                     'Null retention means no mature cohort. Premium interest is not a purchase.',
                     'Partial logs, cleared storage or lost ephemeral disk undercount activity. Usage excludes unreported provider charges.'],
            'observed_events':len(events),'excluded_demo_browsers':len(excluded),
            'pilot_enrolled_browsers':len(enrolled),'activated_browsers':len(users('turn_completed')),
            'completed_browsers':len(users('adventure_completed')),
            'activation_rate':round(len(users('turn_completed'))/len(enrolled),4) if enrolled else None,
            'd7':retention(168,192),'d30':retention(720,744),
            'free_completion_7d':{'eligible_browsers':len(completion_mature),
                'completed_browsers':len(completion_returned),
                'rate':round(len(completion_returned)/len(completion_mature),4) if completion_mature else None,
                'pending_browsers':len(enrolled)-len(completion_mature)},
            'premium_offer_browsers':len(users('premium_viewed')),
            'premium_interest_latest':dict(Counter(interest.values())),
            'feedback_latest':dict(Counter(v['reason'] for v in feedback.values())),
            'feedback_mean_rating':round(sum(v['rating'] for v in feedback.values())/len(feedback),2) if feedback else None,
            'source_enrolled_browsers':{s:len({e['visitor_id'] for e in pilot if e['name']=='enrolled' and e['source']==s})
                                       for s in sorted({e['source'] for e in pilot})},
            'paid_customers':None,'measured_cac':None,'measured_ltv':None,'ai_usage_reported':usage_summary}


if __name__ == '__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('input', nargs='?', help='Local .sqlite3 database or downloaded JSONL/Render log file')
    p.add_argument('--postgres', action='store_true', help='Read the database in ANALYTICS_DATABASE_URL')
    p.add_argument('--as-of', help='ISO timestamp, e.g. 2026-12-01T00:00:00+00:00')
    args=p.parse_args()
    now=datetime.fromisoformat(args.as_of).timestamp() if args.as_of else datetime.now(timezone.utc).timestamp()
    if args.postgres == bool(args.input):
        p.error('Provide one input file or --postgres.')
    if args.postgres:
        from dotenv import load_dotenv
        import psycopg

        load_dotenv(Path(__file__).resolve().parents[1] / '.env')
        database_url = os.getenv('ANALYTICS_DATABASE_URL', '').strip()
        if not database_url:
            p.error('Set ANALYTICS_DATABASE_URL in the environment or tabescape_app/.env.')
        try:
            report = summarize(read_postgres(database_url), now)
        except psycopg.Error:
            p.error('Could not read research PostgreSQL. Check the connection and schema.')
    else:
        report = summarize(read_records(args.input), now)
    print(json.dumps(report, indent=2, ensure_ascii=False))
