import { createClient } from '@supabase/supabase-js';

export default async function handler(req, res) {
  if (req.method !== 'GET') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);
    const { data, error } = await supabase
      .from('rpb_history')
      .select('*')
      .order('created_at', { ascending: false })
      .limit(30);
    if (error) throw error;
    res.status(200).json({ history: data });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
}
