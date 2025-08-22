// api/store.js
import { put } from '@vercel/blob';

export const runtime = 'nodejs';

export default async function handler(req) {
    const t0 = Date.now();
    try {
        if (req.method !== 'POST') {
            return new Response('Method Not Allowed', { status: 405, headers: { Allow: 'POST' } });
        }

        console.log('[store] start', new Date().toISOString());
        const hasToken = !!process.env.BLOB_READ_WRITE_TOKEN;
        console.log('[store] BLOB_READ_WRITE_TOKEN present?', hasToken);

        if (!hasToken) {
            return new Response(
                JSON.stringify({ error: 'Blob store not configured: BLOB_READ_WRITE_TOKEN missing (attach Blob to Production and redeploy)' }),
                { status: 500, headers: { 'content-type': 'application/json' } }
            );
        }

        const { filename, content_type, data_base64 } = await req.json();
        if (!filename || !content_type || !data_base64) {
            return new Response(JSON.stringify({ error: 'Missing fields: filename, content_type, data_base64' }), {
                status: 400, headers: { 'content-type': 'application/json' }
            });
        }

        console.log('[store] calling put() …');
        const buf = Buffer.from(data_base64, 'base64');

        // Guard: time out put() after 10s so we get a clear error instead of a 60s function timeout
        const putPromise = put(filename, buf, {
            access: 'public',
            contentType: content_type,
            addRandomSuffix: true,
            token: process.env.BLOB_READ_WRITE_TOKEN, // pass explicitly to remove ambiguity
        });
        const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error('put() timeout after 10s')), 10_000)
        );
        const blob = await Promise.race([putPromise, timeoutPromise]);

        console.log('[store] put() ok in', Date.now() - t0, 'ms');
        return new Response(JSON.stringify({ url: blob.url }), {
            status: 200, headers: { 'content-type': 'application/json' }
        });
    } catch (e) {
        console.error('[store] error', e);
        return new Response(JSON.stringify({ error: String(e) }), {
            status: 500, headers: { 'content-type': 'application/json' }
        });
    }
}
