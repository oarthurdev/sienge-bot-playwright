module.exports = {
  apps: [
    {
      name: "sienge-autorizar-parcelas",
      script: "./authorize.js",
      interpreter: "node",
      cwd: ".",
      autorestart: true,
      watch: false,
      max_restarts: 20,
      min_uptime: "10s",
      exp_backoff_restart_delay: 2000,
      kill_timeout: 15000,
      env: {
        NODE_ENV: "production",
        PM2_LOOP: "true",
        PM2_INTERVAL_MS: "60000",
        PM2_STOP_ON_FATAL: "false",
        HEADLESS: "true",
        DEBUG_HTML: "false"
      }
    }
  ]
};
