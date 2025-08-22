import { put } from '@vercel/blob';

export const runtime = 'nodejs';

export default async function handler(req, res) {
    const t0 = Date.now();
    try {
        if (req.method !== 'POST') {
            res.setHeader('Allow', 'POST');
            return res.status(405).end('Method Not Allowed');
        }

        console.log('[store] start', new Date().toISOString());

        const hasToken = !!process.env.BLOB_READ_WRITE_TOKEN;
        console.log('[store] BLOB_READ_WRITE_TOKEN present?', hasToken);
        if (!hasToken) {
            return res.status(500).json({
                error:
                    'Blob store not configured: BLOB_READ_WRITE_TOKEN missing. Attach Blob to this environment and redeploy.',
            });
        }

        // Parse JSON body (works for raw Node streams and frameworks)
        let payload = req.body;
        if (!payload || typeof payload !== 'object') {
            const chunks = [];
            for await (const chunk of req) {
                chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
            }
            const raw = Buffer.concat(chunks).toString('utf8');
            payload = raw ? JSON.parse(raw) : {};
        }

        const { filename, content_type, data_base64 } = payload || {};
        if (!filename || !content_type || !data_base64) {
            return res
                .status(400)
                .json({ error: 'Missing fields: filename, content_type, data_base64' });
        }

        console.log('[store] calling put() …');
        const buf = Buffer.from(data_base64, 'base64');

        // Don’t hang forever if network hiccups — bail out after 10s
        const putPromise = put(filename, buf, {
            access: 'public',
            contentType: content_type,
            addRandomSuffix: true,
            token: process.env.BLOB_READ_WRITE_TOKEN,
        });
        const timeoutPromise = new Promise((_, reject) =>
            setTimeout(() => reject(new Error('put() timeout after 10s')), 10_000)
        );
        const blob = await Promise.race([putPromise, timeoutPromise]);

        console.log('[store] put() ok in', Date.now() - t0, 'ms');
        return res.status(200).json({ url: blob.url });
    } catch (e) {
        console.error('[store] error', e);
        return res.status(500).json({ error: String(e?.message || e) });
    }
}
