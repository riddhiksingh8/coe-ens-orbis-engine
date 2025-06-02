import pkg from 'pg';
import dotenv from 'dotenv';
import { emitSessionStatus } from '../utils/socketUtil.js';

dotenv.config();

const { Pool, Client } = pkg;

const config = {
  user: process.env.DB_USER,
  host: process.env.DB_HOST,
  database: process.env.DB_NAME,
  password: process.env.DB_PASSWORD,
  port: Number(process.env.DB_PORT),
  ssl: { rejectUnauthorized: false },
};

// --- Create Pool for general queries ---
const pool = new Pool(config);

// Handle pool events
pool.on('connect', () => {
  console.log('Connection pool established with the database');
});

pool.on('error', (err) => {
  console.error('Connection pool error:', err);
});

// --- Create a dedicated client for LISTEN/NOTIFY ---
const listenerClient = new Client(config);

async function setupListener() {
  try {
    await listenerClient.connect();
    console.log('Listener client connected to database');

    await listenerClient.query('LISTEN session_id_status_channel');

    listenerClient.on('notification', (msg) => {
      try {
        const payload = JSON.parse(msg.payload);

        emitSessionStatus(payload);
      } catch (err) {
        console.error('Failed to parse payload:', err);
      }
    });

    listenerClient.on('error', (err) => {
      console.error('Listener client error:', err);
    });
  } catch (err) {
    console.error('Failed to set up listener:', err);
  }
}

// Initialize listener
setupListener();

// Export pool for use in queries
export default pool;
