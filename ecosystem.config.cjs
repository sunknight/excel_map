// PM2 配置：Excel 脑图生成工具（静态站点，http-server 托管）
// 启动:  pm2 start ecosystem.config.cjs
// 重启:  pm2 reload ecosystem.config.cjs   （改配置/更新文件后必须带文件名）
// 停止:  pm2 stop excel-map
// 删除:  pm2 delete excel-map
module.exports = {
  apps: [
    {
      name: 'excel-map',
      script: 'node_modules/http-server/bin/http-server',
      args: '-p 38195 -a 0.0.0.0 -c-1 ./',
      cwd: __dirname,
      watch: false,
      autorestart: true,
      max_memory_restart: '200M',
      env: {
        NODE_ENV: 'production',
      },
    },
  ],
};
