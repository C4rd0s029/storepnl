const stripe = require('stripe')(process.env.STRIPE_SECRET_KEY);
const { createClient } = require('@supabase/supabase-js');

const supabase = createClient(
  process.env.REACT_APP_SUPABASE_URL,
  process.env.SUPABASE_SERVICE_KEY
);

export const config = { api: { bodyParser: false } };

async function buffer(readable) {
  const chunks = [];
  for await (const chunk of readable) chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : chunk);
  return Buffer.concat(chunks);
}

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).end();

  const sig = req.headers['stripe-signature'];
  const buf = await buffer(req);

  let event;
  try {
    event = stripe.webhooks.constructEvent(buf, sig, process.env.STRIPE_WEBHOOK_SECRET);
  } catch (err) {
    return res.status(400).json({ error: `Webhook Error: ${err.message}` });
  }

  const userId = event.data.object?.metadata?.userId || event.data.object?.subscription_data?.metadata?.userId;

  if (event.type === 'checkout.session.completed' || event.type === 'invoice.payment_succeeded') {
    if (userId) {
      await supabase.from('profiles').upsert({ id: userId, plan: 'pro', plan_updated_at: new Date().toISOString() });
    }
  }

  if (event.type === 'customer.subscription.deleted') {
    if (userId) {
      await supabase.from('profiles').upsert({ id: userId, plan: 'free', plan_updated_at: new Date().toISOString() });
    }
  }

  res.status(200).json({ received: true });
};
