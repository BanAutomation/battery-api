import { put } from '@vercel/blob';

export const runtime = 'nodejs'; // ok to omit; Node runtime is default for .ts

export default async function handler(req: Request): Promise<Response> {
  if (req.method !== 'POST') {
    return new Response('Method Not Allowed', { status: 405, headers: { Allow: 'POST' } });
  }
  try {
    const { filename, content_type, data_base64 } = await req.json() as any;

    if (!filename || !content_type || !data_base64) {
      return new Response(
        JSON.stringify({ error: 'Missing fields: filename, content_type, data_base64' }),
        { status: 400, headers: { 'content-type': 'application/json' } }
      );
    }

    const buf = Buffer.from(data_base64, 'base64');
    const blob = await put(filename, buf, {
      access: 'public',
      contentType: content_type,
      addRandomSuffix: true, // fights CDN caching/overwrites
    });

    return new Response(JSON.stringify({ url: blob.url }), {
      status: 200,
      headers: { 'content-type': 'application/json' },
    });
  } catch (e: any) {
    return new Response(JSON.stringify({ error: String(e) }), {
      status: 500,
      headers: { 'content-type': 'application/json' },
    });
  }
}