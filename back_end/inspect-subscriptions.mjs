import pkg from 'pg';
import dotenv from 'dotenv';

dotenv.config();
const { Client } = pkg;
const client = new Client({ connectionString: process.env.DATABASE_URL });
try {
  await client.connect();
  const cols = await client.query("select column_name, data_type, is_nullable, column_default from information_schema.columns where table_name='subscriptions' order by ordinal_position");
  console.log('columns:', cols.rows);
  const cons = await client.query("select conname, contype, pg_get_constraintdef(oid) as def from pg_constraint where conrelid='subscriptions'::regclass");
  console.log('constraints:', cons.rows);
  const idx = await client.query("select indexname, indexdef from pg_indexes where tablename='subscriptions'");
  console.log('indexes:', idx.rows);
  const rows = await client.query('select * from subscriptions limit 5');
  console.log('sample rows:', rows.rows);
} finally {
  await client.end();
}
