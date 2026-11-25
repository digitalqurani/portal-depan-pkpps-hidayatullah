PRAGMA foreign_keys = ON;
CREATE TABLE IF NOT EXISTS video_jobs (
 id TEXT PRIMARY KEY, organization_id TEXT NOT NULL, owner_portal_user_id TEXT NOT NULL,
 provider TEXT NOT NULL DEFAULT 'fal', model_id TEXT NOT NULL, provider_request_id TEXT,
 status TEXT NOT NULL DEFAULT 'queued' CHECK(status IN ('queued','processing','completed','failed','cancelled')),
 prompt TEXT NOT NULL, duration_seconds INTEGER NOT NULL, aspect_ratio TEXT NOT NULL,
 provider_status_url TEXT, provider_response_url TEXT, output_url TEXT, error_message TEXT,
 created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_video_jobs_owner ON video_jobs(organization_id,owner_portal_user_id,created_at DESC);
CREATE INDEX IF NOT EXISTS idx_video_jobs_request ON video_jobs(provider,provider_request_id);
CREATE TABLE IF NOT EXISTS video_projects (
 id TEXT PRIMARY KEY, organization_id TEXT NOT NULL, owner_portal_user_id TEXT NOT NULL,
 title TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'draft',
 created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS video_assets (
 id TEXT PRIMARY KEY, organization_id TEXT NOT NULL, job_id TEXT NOT NULL,
 asset_type TEXT NOT NULL, storage_key TEXT NOT NULL, mime_type TEXT, size_bytes INTEGER,
 created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, FOREIGN KEY(job_id) REFERENCES video_jobs(id)
);
CREATE TABLE IF NOT EXISTS video_wallets (
 organization_id TEXT NOT NULL, portal_user_id TEXT NOT NULL,
 balance_credits INTEGER NOT NULL DEFAULT 0 CHECK(balance_credits>=0),
 updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, PRIMARY KEY(organization_id,portal_user_id)
);
CREATE TABLE IF NOT EXISTS video_credit_transactions (
 id TEXT PRIMARY KEY, organization_id TEXT NOT NULL, portal_user_id TEXT NOT NULL,
 job_id TEXT, transaction_type TEXT NOT NULL, amount_credits INTEGER NOT NULL,
 idempotency_key TEXT NOT NULL UNIQUE, note TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS video_model_catalog (
 model_id TEXT PRIMARY KEY, provider TEXT NOT NULL, display_name TEXT NOT NULL,
 enabled INTEGER NOT NULL DEFAULT 0, input_schema_json TEXT, cost_reference_json TEXT
);
CREATE TABLE IF NOT EXISTS video_audit_logs (
 id TEXT PRIMARY KEY, organization_id TEXT NOT NULL, portal_user_id TEXT,
 action TEXT NOT NULL, resource_type TEXT, resource_id TEXT, metadata_json TEXT,
 created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
