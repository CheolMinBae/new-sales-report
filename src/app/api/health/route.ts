import { NextResponse } from 'next/server';
import { getS3ConfigInfo } from '@/lib/dataStore';

export async function GET() {
  try {
    const s3Config = await getS3ConfigInfo();
    
    return NextResponse.json({
      status: 'healthy',
      timestamp: new Date().toISOString(),
      uptime: process.uptime(),
      environment: process.env.NODE_ENV || 'development',
      s3: {
        configured: s3Config.hasCredentials,
        bucket: s3Config.bucketName,
        region: s3Config.region,
      },
      version: process.env.npm_package_version || '1.0.0',
    });
  } catch (error) {
    console.error('Health check error:', error);
    return NextResponse.json({
      status: 'unhealthy',
      timestamp: new Date().toISOString(),
      error: error instanceof Error ? error.message : 'Unknown error',
    }, { status: 500 });
  }
} 