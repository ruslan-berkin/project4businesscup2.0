CREATE TABLE IF NOT EXISTS research_events (
    dedupe TEXT PRIMARY KEY,
    ts BIGINT NOT NULL,
    payload TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS research_events_time ON research_events (ts, dedupe);
