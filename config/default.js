module.exports = {
  // 服务器配置
  server: {
    port: process.env.PORT || 3000,
    host: process.env.HOST || 'localhost'
  },
  
  // 文件上传配置
  upload: {
    maxFileSize: 50 * 1024 * 1024, // 50MB
    allowedTypes: [
      'application/pdf',
      'application/msword',
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'application/vnd.ms-excel',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-powerpoint',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'text/csv',
      'text/markdown',
      'text/plain'
    ],
    uploadDir: './uploads'
  },
  
  // 文件预览配置
  preview: {
    cacheDir: './cache',
    maxCacheSize: 100 * 1024 * 1024, // 100MB
    cacheExpire: 24 * 60 * 60 * 1000 // 24小时
  },
  
  // 安全配置
  security: {
    cors: {
      origin: process.env.CORS_ORIGIN || '*',
      credentials: true
    },
    helmet: {
      contentSecurityPolicy: {
        directives: {
          defaultSrc: ["'self'"],
          styleSrc: ["'self'", "'unsafe-inline'", "https://oss.vtopideal.com"],
          scriptSrc: ["'self'", "https://oss.vtopideal.com"],
          imgSrc: ["'self'", "data:", "https:"],
          connectSrc: ["'self'"]
        }
      }
    }
  },
  
  // 日志配置
  logging: {
    level: process.env.LOG_LEVEL || 'info',
    filename: './logs/app.log',
    maxSize: '10m',
    maxFiles: '5'
  }
}; 