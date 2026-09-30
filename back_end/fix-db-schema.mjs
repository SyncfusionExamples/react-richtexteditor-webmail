import pkg from 'pg';
const { Client } = pkg;
const client = new Client({ connectionString: process.env.DATABASE_URL || 'postgres://postgres:postgres123@localhost:5432/maildb' });
try {
  await client.connect();
  await client.query("ALTER TABLE subscriptions ADD COLUMN IF NOT EXISTS delta_link text");
  const res = await client.query("SELECT column_name, data_type FROM information_schema.columns WHERE table_name='subscriptions' ORDER BY ordinal_position");
  console.log(res.rows);
} finally {
  await client.end();
}
