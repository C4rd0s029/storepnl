const stripe = require('stripe')(process.env.STRIPE_SECRET_KEY);

module.exports = async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const { userId, email, fromLanding } = req.body;

  try {
    const sessionParams = {
      mode: 'subscription',
      payment_method_types: ['card'],
      line_items: [{
        price: process.env.REACT_APP_STRIPE_PRICE_ID,
        quantity: 1,
      }],
      success_url: `${req.headers.origin}/?payment=success`,
      cancel_url: `${req.headers.origin}/`,
    };

    // From landing — no user yet, collect email in Stripe
    if (fromLanding) {
      sessionParams.success_url = `${req.headers.origin}/?payment=success`;
    } else {
      // From settings — user already exists
      sessionParams.customer_email = email;
      sessionParams.metadata = { userId };
      sessionParams.subscription_data = { metadata: { userId } };
      sessionParams.success_url = `${req.headers.origin}/?payment=success&userId=${userId}`;
    }

    const session = await stripe.checkout.sessions.create(sessionParams);
    res.status(200).json({ url: session.url });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};
