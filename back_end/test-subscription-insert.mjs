import pkg from 'pg';
import dotenv from 'dotenv';

dotenv.config();
const { Client } = pkg;
const client = new Client({ connectionString: process.env.DATABASE_URL });
try {
  await client.connect();
  const testId = 'test-subscription-0001';
  await client.query('DELETE FROM subscriptions WHERE id = $1 OR subscription_id = $1', [testId]);
  await client.query(
    'INSERT INTO subscriptions (id, subscription_id, user_id, resource, change_type, notification_url, expiration_date_time, client_state, delta_link) VALUES ($1,$1,$2,$3,$4,$5,$6,$7,$8)',
    [testId, 'user_test', "me/mailFolders('Inbox')/messages", 'created,updated,deleted', 'https://example.com/webhook', new Date().toISOString(), 'steed-client', null]
  );
  const row = await client.query('SELECT * FROM subscriptions WHERE id = $1', [testId]);
  console.log(row.rows);
  await client.query('DELETE FROM subscriptions WHERE id = $1', [testId]);
} finally {
  await client.end();
}
