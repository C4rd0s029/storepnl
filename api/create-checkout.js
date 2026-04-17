const stripe = require('stripe')(process.env.STRIPE_SECRET_KEY);

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { userId, email } = req.body;

  if (!userId || !email) {
    return res.status(400).json({ error: 'Missing userId or email' });
  }

  try {
    const session = await stripe.checkout.sessions.create({
      mode: 'subscription',
      payment_method_types: ['card'],
      customer_email: email,
      line_items: [{
        price: process.env.REACT_APP_STRIPE_PRICE_ID,
        quantity: 1,
      }],
      success_url: `${req.headers.origin}/?payment=success&userId=${userId}`,
      cancel_url: `${req.headers.origin}/?payment=cancelled`,
      metadata: { userId },
      subscription_data: { metadata: { userId } },
    });

    res.status(200).json({ url: session.url });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
