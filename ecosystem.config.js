/**
 * ecosystem.config.js  –  PM2 configuration for QuoToCon
 *
 * Production deployment on Ubuntu:
 *
 *   1. Activate the project virtual-env so gunicorn is on PATH, then run:
 *        pm2 start ecosystem.config.js --env production
 *
 *   2. Persist across reboots:
 *        pm2 save && pm2 startup
 *
 * Environment variables that MUST be set in production (via a .env file or
 * shell export before pm2 start):
 *   SECRET_KEY  – a long, random string used to sign Flask sessions
 *   PORT        – (optional) listening port, default 5000
 */
module.exports = {
  apps: [
    {
      name: 'QuoToCon',

      // gunicorn is the production WSGI server; point at the Flask app object.
      // Adjust the path to gunicorn if you use a virtualenv in a non-standard
      // location (e.g. /home/ubuntu/QuoToCon/venv/bin/gunicorn).
      script: 'gunicorn',
      args: '-w 2 -b 0.0.0.0:5001 --timeout 120 --access-logfile - app:app',
      interpreter: 'none',

      // Working directory – all relative paths in app.py resolve from here
      cwd: '/home/ubuntu/QuoToCon',

      instances: 1,
      autorestart: true,
      watch: false,
      max_memory_restart: '300M',

      env: {
        FLASK_ENV: 'development',
        PORT: '5001',
      },
      env_production: {
        FLASK_ENV: 'production',
        PORT: '5001',
        // SECRET_KEY should be injected from the host environment or a .env
        // file; never hard-code it here.
      },

      log_date_format: 'YYYY-MM-DD HH:mm:ss Z',
      error_file: '/home/ubuntu/QuoToCon/logs/err.log',
      out_file: '/home/ubuntu/QuoToCon/logs/out.log',
      merge_logs: true,
    },
  ],
};
