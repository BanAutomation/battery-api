// api/store.ts
import { NextRequest, NextResponse } from 'next/server';
import { put } from '@vercel/blob';

export const runtime = 'nodejs';

export async function POST(req: NextRequest) {
  try {
    const { filename, content_type, data_base64 } = await req.json();
    if (!filename || !content_type || !data_base64) {
      return NextResponse.json({ error: 'Missing fields' }, { status: 400 });
    }
    const buf = Buffer.from(data_base64, 'base64');
    const blob = await put(filename, buf, { access: 'public', contentType: content_type });
    return NextResponse.json({ url: blob.url });
  } catch (e: any) {
    return NextResponse.json({ error: String(e) }, { status: 500 });
  }
}
