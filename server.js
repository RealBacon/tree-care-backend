const express = require('express');
const Stripe = require('stripe');
const { Client } = require('@microsoft/microsoft-graph-client');
const { BlobServiceClient } = require('@azure/storage-blob');
const cors = require('cors');
const fetch = require('isomorphic-fetch');

const app = express();
const stripe = Stripe(process.env.STRIPE_SECRET_KEY);
const azureConnectionString = process.env.AZURE_STORAGE_CONNECTION_STRING;
const blobServiceClient = BlobServiceClient.fromConnectionString(azureConnectionString);
const containerClient = blobServiceClient.getContainerClient('photos');

app.use(cors());
app.use(express.json());
app.use(express.raw({ type: 'multipart/form-data', limit: '10mb' }));

const graphClient = Client.init({
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
            const token = await response.json();
            done(null, token.access_token);
        } catch (error) {
            done(error);
        }
    }
});

app.get('/events', async (req, res) => {
    try {
        const events = await graphClient.api('/users/support@propertreecare.com/calendar/events').get();
        res.json(events.value);
    } catch (error) {
        console.error('Error fetching events:', error);
        res.status(500).json({ error: 'Failed to fetch events' });
    }
});

app.post('/check-availability', async (req, res) => {
    const { startTime, endTime } = req.body;
    try {
        const events = await graphClient.api('/users/support@propertreecare.com/calendar/events')
            .filter(`start/dateTime ge '${startTime}' and end/dateTime le '${endTime}'`)
            .get();
        if (events.value.length > 0) {
            res.status(400).json({ error: 'Time slot is booked' });
        } else {
            res.json({ available: true });
        }
    } catch (error) {
        console.error('Error checking availability:', error);
        res.status(500).json({ error: 'Failed to check availability' });
    }
});

app.post('/upload-photos', async (req, res) => {
    try {
        const photoUrls = [];
        for (const photo of req.files.photos) {
            const blobName = `photo-${Date.now()}-${photo.name}`;
            const blockBlobClient = containerClient.getBlockBlobClient(blobName);
            await blockBlobClient.uploadData(photo.data);
            photoUrls.push(blockBlobClient.url);
        }
        res.json(photoUrls);
    } catch (error) {
        console.error('Error uploading photos:', error);
        res.status(500).json({ error: 'Failed to upload photos' });
    }
});

app.post('/create-checkout-session', async (req, res) => {
    try {
        const { name, email, phone, duration, price, startTime, endTime, timezone, notes, photoUrls } = req.body;

        await graphClient.api('/users/support@propertreecare.com/calendar/events').post({
            subject: `Consultation with ${name}`,
            body: {
                contentType: 'HTML',
                content: `Client: ${name}<br>Email: ${email}<br>Phone: ${phone}<br>Notes: ${notes}<br>Photos: ${photoUrls.join('<br>')}`
            },
            start: { dateTime: startTime, timeZone: timezone },
            end: { dateTime: endTime, timeZone: timezone },
            attendees: [{ emailAddress: { address: email, name } }]
        });

        const session = await stripe.checkout.sessions.create({
            payment_method_types: ['card'],
            line_items: [{
                price_data: {
                    currency: 'usd',
                    product_data: {
                        name: `Virtual Arborist Consultation (${duration} minutes)`,
                        metadata: { name, email, phone, startTime, notes }
                    },
                    unit_amount: price
                },
                quantity: 1
            }],
            mode: 'payment',
            success_url: 'https://propertreecare.com/consultation.html?success=true',
            cancel_url: 'https://propertreecare.com/consultation.html',
            customer_email: email
        });

        await graphClient.api('/users/support@propertreecare.com/messages').post({
            subject: 'Consultation Booking Confirmation',
            body: {
                contentType: 'HTML',
                content: `Dear ${name},<br>Your ${duration}-minute consultation is confirmed for ${moment(startTime).format('MMMM D, YYYY h:mm A')} (${timezone}).<br>Notes: ${notes}<br>Photos: ${photoUrls.join('<br>')}<br>Thank you for choosing Proper Tree Care!`
            },
            toRecipients: [{ emailAddress: { address: email } }]
        });
        await graphClient.api('/users/support@propertreecare.com/messages').send();

        res.json({ id: session.id });
    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: 'Failed to create checkout session' });
    }
});

app.listen(3000, () => console.log('Server running on port 3000'));