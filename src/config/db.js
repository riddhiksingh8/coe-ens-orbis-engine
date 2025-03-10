import pkg from "pg";
import dotenv from "dotenv";

const { Pool } = pkg;
dotenv.config();

// Configure the pool with SSL
const pool = new Pool({
  user: process.env.DB_USER,
  host: process.env.DB_HOST,
  database: process.env.DB_NAME,
  password: process.env.DB_PASSWORD,
  port: process.env.DB_PORT,
  ssl: {
    rejectUnauthorized: false, // Use true for production with valid certificates
  },
});

// Event listener for pool connection
pool.on("connect", () => {
  console.log("Connection pool established with Database");
});
// Event listener for pool connection
pool.on("error", (error) => {
  console.log("Connection pool Failed", error);
});
// Export the pool instance
export default pool;