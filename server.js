const express = require('express');
const cors = require('cors');
const stripe = require('stripe')(process.env.STRIPE_SECRET_KEY);
const { Client } = require('@microsoft/microsoft-graph-client');
const { BlobServiceClient } = require('@azure/storage-blob');
const multer = require('multer');
const moment = require('moment');
const app = express();

// Middleware
app.use(cors());
app.use(express.json());
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } }); // 10MB limit

// Azure Blob Storage
const azureConnectionString = process.env.AZURE_STORAGE_CONNECTION_STRING;
let containerClient;
if (azureConnectionString) {
    const blobServiceClient = BlobServiceClient.fromConnectionString(azureConnectionString);
    containerClient = blobServiceClient.getContainerClient('photos');
}

// Microsoft Graph Client
let graphClient;
if (process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET) {
    graphClient = Client.init({
        authProvider: async (done) => {
            try {
                const response = await fetch(`https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams({
                        client_id: process.env.CLIENT_ID,
                        client_secret: process.env.CLIENT_SECRET,
                        scope: 'https://graph.microsoft.com/.default',
                        grant_type: 'client_credentials'
                    })
                });
                if (!response.ok) {
                    throw new Error(`Token request failed: ${response.statusText}`);
                }
                const token = await response.json();
                done(null, token.access_token);
            } catch (error) {
                console.error('Graph auth error:', error);
                done(error);
            }
        }
    });
}

// Get calendar events
app.get('/events', async (req, res) => {
    try {
        if (!graphClient) {
            return res.status(503).json({ error: 'Microsoft Graph not configured' });
        }
        const events = await graphClient.api('/users/support@propertreecare.com/calendar/events').get();
        res.json(events.value || []);
    } catch (error) {
        console.error('Error fetching events:', error.message);
        res.status(500).json({ error: 'Failed to fetch events' });
    }
});

// Check availability
app.post('/check-availability', async (req, res) => {
    const { startTime, endTime } = req.body;
    if (!startTime || !endTime) {
        return res.status(400).json({ error: 'Missing startTime or endTime' });
    }
    try {
        if (!graphClient) {
            return res.status(503).json({ error: 'Microsoft Graph not configured' });
        }
        const events = await graphClient.api('/users/support@propertreecare.com/calendarview')
            .query({ startDateTime: startTime, endDateTime: endTime })
            .get();
        if (events.value.length > 0) {
            return res.status(400).json({ error: 'Time slot is booked' });
        }
        res.json({ available: true });
    } catch (error) {
        console.error('Error checking availability:', error.message);
        res.status(500).json({ error: 'Failed to check availability' });
    }
});

// Upload photos
app.post('/upload-photos', upload.array('photos', 5), async (req, res) => {
    try {
        if (!containerClient) {
            return res.status(503).json({ error: 'Azure Blob Storage not configured' });
        }
        const photoUrls = [];
        if (req.files && req.files.length > 0) {
            for (const file of req.files) {
                const blobName = `photo-${Date.now()}-${file.originalname}`;
                const blockBlobClient = containerClient.getBlockBlobClient(blobName);
                await blockBlobClient.uploadData(file.buffer);
                photoUrls.push(blockBlobClient.url);
            }
        }
        res.json(photoUrls);
    } catch (error) {
        console.error('Photo upload error:', error.message);
        res.status(500).json({ error: 'Failed to upload photos' });
    }
});

// Create Stripe Checkout session
app.post('/create-checkout-session', async (req, res) => {
    try {
        const { name, email, phone, duration, price, startTime, endTime, timezone, notes, photoUrls } = req.body;
        if (!name || !email || !duration || !price || !startTime || !endTime || !timezone) {
            return res.status(400).json({ error: 'Missing required fields' });
        }

        // Create calendar event
        if (graphClient) {
            await graphClient.api('/users/support@propertreecare.com/calendar/events').post({
                subject: `Consultation with ${name}`,
                body: {
                    contentType: 'HTML',
                    content: `Client: ${name}<br>Email: ${email}<br>Phone: ${phone}<br>Notes: ${notes || 'None'}<br>Photos: ${(photoUrls || []).join('<br>') || 'None'}`
                },
                start: { dateTime: startTime, timeZone: timezone },
                end: { dateTime: endTime, timeZone: timezone },
                attendees: [{ emailAddress: { address: email, name } }]
            });
        }

        // Create Stripe Checkout session
        const session = await stripe.checkout.sessions.create({
            payment_method_types: ['card'],
            line_items: [{
                price_data: {
                    currency: 'usd',
                    product_data: {
                        name: `Virtual Arborist Consultation (${duration} minutes)`,
                        description: 'Consultation with Proper Tree Care',
                        metadata: { name, email, phone, startTime, notes }
                    },
                    unit_amount: parseInt(price) // Ensure price is an integer (cents)
                },
                quantity: 1
            }],
            mode: 'payment',
            success_url: 'https://propertreecare.com/success.html',
            cancel_url: 'https://propertreecare.com/consultation.html',
            customer_email: email
        });

        // Send confirmation email
        if (graphClient) {
            const message = {
                subject: 'Consultation Booking Confirmation',
                body: {
                    contentType: 'HTML',
                    content: `Dear ${name},<br>Your ${duration}-minute consultation is confirmed for ${moment(startTime).format('MMMM D, YYYY h:mm A')} (${timezone}).<br>Notes: ${notes || 'None'}<br>Photos: ${(photoUrls || []).join('<br>') || 'None'}<br>Thank you for choosing Proper Tree Care!`
                },
                toRecipients: [{ emailAddress: { address: email } }]
            };
            await graphClient.api('/users/support@propertreecare.com/messages').post(message);
            await graphClient.api(`/users/support@propertreecare.com/messages/${message.id}/send`).post();
        }

        res.json({ id: session.id });
    } catch (error) {
        console.error('Checkout session error:', error.message);
        res.status(500).json({ error: 'Failed to create checkout session' });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));