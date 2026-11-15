import contextlib
import io
import json
import os
import sqlite3
import tempfile
import unittest
import uuid
from unittest.mock import patch
import app
from scripts.analytics_report import summarize


class ResearchTests(unittest.TestCase):
    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.db=os.path.join(self.temp.name,'events.sqlite3')
        self.env=patch.dict(os.environ,{'ANALYTICS_DB_PATH':self.db,'ANALYTICS_DATABASE_URL':'','AI_ENABLED':'false'})
        self.env.start(); self.addCleanup(self.env.stop);self.addCleanup(self.temp.cleanup)
        app._RESEARCH_WINDOW.clear()
        self.client=app.app.test_client()
        self.payload=dict(event_id=str(uuid.uuid4()),visitor_id=str(uuid.uuid4()),consent=True,
                          name='enrolled',cohort='pilot',source='direct',eligible=True,workspace='excel')
    def send(self, **changes):
        with contextlib.redirect_stdout(io.StringIO()):
            return self.client.post('/api/research/event',json=dict(self.payload,**changes))
    def test_consent_schema_origin_and_no_sensitive_fields(self):
        self.assertEqual(self.send(consent=False).status_code,400)
        self.assertEqual(self.send(action='private text').status_code,400)
        self.assertEqual(self.send(name=[]).status_code,400)
        self.assertEqual(self.client.post('/api/research/event',json=self.payload,headers={'Origin':'https://other.example'}).status_code,403)
        self.assertEqual(self.send().status_code,202)
        self.assertEqual(self.send().json['duplicate'],True)
        with sqlite3.connect(self.db) as db: record=json.loads(db.execute('SELECT payload FROM events').fetchone()[0])
        self.assertNotIn('story_token',record);self.assertNotIn('ip',record)
    def test_verified_game_events_deduplicate_and_reject_fake_completion(self):
        token=self.client.post('/api/start',json={'campaign_id':'oakhaven','mode':'free','language':'en'}).json['story_token']
        kwargs=dict(name='adventure_started',adventure_id=str(uuid.uuid4()),story_token=token)
        self.assertEqual(self.send(**kwargs).status_code,202)
        self.assertTrue(self.send(**kwargs,event_id=str(uuid.uuid4())).json['duplicate'])
        self.assertEqual(self.send(**dict(kwargs,name='adventure_completed')).status_code,400)
        self.assertEqual(self.send(**dict(kwargs,story_token='broken')).status_code,400)
    def test_storage_failure_never_blocks_game(self):
        with patch('sqlite3.connect',side_effect=sqlite3.OperationalError('unavailable')):
            self.assertEqual(self.send().status_code,202)
        self.assertEqual(self.client.get('/healthz').status_code,200)
    def test_postgres_failure_does_not_confirm_a_lost_event(self):
        import psycopg

        with self.assertLogs(app.app.logger, level="WARNING") as logs, \
                patch.dict(os.environ, {"ANALYTICS_DATABASE_URL": "postgresql://test"}), \
                patch("psycopg.connect", side_effect=psycopg.OperationalError("private connection details")):
            response = self.send()
        self.assertEqual(response.status_code, 503)
        self.assertNotIn("private connection details", response.get_data(as_text=True))
        self.assertNotIn("private connection details", str(logs.output))
        self.assertFalse(os.path.exists(self.db))
        self.assertEqual(self.client.get("/healthz").status_code, 200)
    def test_intake_rate_limit(self):
        app._RESEARCH_WINDOW[:]=[app.time.time()]*300
        self.assertEqual(self.send().status_code,429)
    def test_retention_mature_windows_demo_exclusion_and_deduplication(self):
        base=1700000000
        def event(v,n,h,**kw):
            return dict(event_id=str(uuid.uuid4()),visitor_id=v,name=n,ts=base+h*3600,
                        cohort='pilot',source='direct',eligible=True,**kw)
        events=[event('a','enrolled',0),event('a','turn_completed',1,adventure_id='one',turn=1),
                event('a','turn_completed',170,adventure_id='two',turn=1),
                event('b','enrolled',150),event('b','turn_completed',151,adventure_id='three',turn=1),
                event('jury','enrolled',0),dict(event('jury','turn_completed',1),cohort='demo')]
        events.append(event('a','adventure_completed',5,adventure_id='one',turn=6,mode='free'))
        events.append(event('b','adventure_completed',152,adventure_id='three',turn=6,mode='free'))
        events.append(dict(events[1],event_id=str(uuid.uuid4()),ts=base+171*3600))
        result=summarize(events,base+200*3600)
        self.assertEqual(result['pilot_enrolled_browsers'],2)
        self.assertEqual(result['d7'],dict(eligible_browsers=1,returned_browsers=1,rate=1.0,pending_browsers=1))
        self.assertEqual(result['free_completion_7d'],dict(eligible_browsers=1,completed_browsers=1,rate=1.0,pending_browsers=1))
        self.assertIsNone(result['d30']['rate']);self.assertIsNone(result['measured_cac'])
    def test_provider_usage_missing_is_not_zero_and_gemini_includes_thinking(self):
        out=io.StringIO()
        with contextlib.redirect_stdout(out):
            app.log_ai_usage('gemini','test',{'usageMetadata':{'promptTokenCount':10,'candidatesTokenCount':20,'thoughtsTokenCount':5}})
            app.log_ai_usage('groq','test',{})
        records=[json.loads(x.split(' ',1)[1]) for x in out.getvalue().splitlines()]
        self.assertEqual(records[0]['output_tokens'],25)
        self.assertIsNone(records[1]['input_tokens'])
        result=summarize(records,app.time.time()+1)
        self.assertEqual(result['ai_usage_reported']['groq/test']['usage_missing'],1)
