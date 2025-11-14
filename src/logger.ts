import winston from 'winston';
import path from 'path';
import { fileURLToPath } from 'url';
import fs from 'fs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const logsDir = path.join(__dirname, '..', 'logs');

// Customize colors for console output
// info -> green (existing), debug -> magenta ("purple"), warn -> yellow, error -> red
// Note: Requires enableConsoleLogging() to add the Console transport
winston.addColors({
  info: 'green',
  debug: 'magenta',
  warn: 'yellow',
  error: 'red',
});

if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir);
}

const logger = winston.createLogger({
  level: process.env.MS365_MCP_LOG_LEVEL || 'info',
  format: winston.format.combine(
    winston.format.timestamp({
      format: 'YYYY-MM-DD HH:mm:ss',
    }),
    // Include meta fields (e.g., preview, bytes) in file logs
    winston.format.printf((info) => {
      const { level, message, timestamp, ...meta } = info as any;
      const metaStr =
        meta && Object.keys(meta).length > 0 ? ` ${JSON.stringify(meta, null, 0)}` : '';
      return `${timestamp} ${String(level).toUpperCase()}: ${message}${metaStr}`;
    })
  ),
  transports: [
    new winston.transports.File({
      filename: path.join(logsDir, 'error.log'),
      level: 'error',
    }),
    new winston.transports.File({
      filename: path.join(logsDir, 'mcp-server.log'),
    }),
  ],
});

let consoleTransportAdded = false;
export const enableConsoleLogging = (): void => {
  if (consoleTransportAdded) return;
  logger.add(
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
        // Include meta fields in console logs; color only the level label
        winston.format.printf((info) => {
          const { level, message, timestamp, ...meta } = info as any;
          const colorizer = winston.format.colorize();
          const levelLabel = `${String(level).toLowerCase()}:`;
          const coloredLevel = colorizer.colorize(String(level), levelLabel);
          const metaStr =
            meta && Object.keys(meta).length > 0 ? ` ${JSON.stringify(meta, null, 0)}` : '';
          return `${timestamp} ${coloredLevel} ${message}${metaStr}`;
        })
      ),
      silent: process.env.MS365_MCP_SILENT === 'true' || process.env.MS365_MCP_SILENT === '1',
    })
  );
  // Default to debug when console logging is enabled unless overridden via env
  logger.level = (process.env.MS365_MCP_LOG_LEVEL || 'debug').toLowerCase();
  consoleTransportAdded = true;
};

export default logger;
