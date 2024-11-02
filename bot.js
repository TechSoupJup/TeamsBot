const { ActivityHandler, MessageFactory, CardFactory } = require('botbuilder');
const axios = require('axios');
const fs = require('fs');
const https = require('https');
require('dotenv').config();

// Configure HTTPS agent with Zscaler certificate
const httpsAgent = new https.Agent({
    ca: fs.readFileSync('./Zscaler Root CA.cer')  // Ensure this matches the exact file name of the Zscaler certificate
});

class FreshserviceBot extends ActivityHandler {
    constructor() {
        super();

        // Respond to incoming messages
        this.onMessage(async (context, next) => {
            const message = context.activity.text.toLowerCase();
            if (message.includes('approved changes')) {
                const changes = await this.getApprovedChanges();
                const reply = this.formatChangesForTeams(changes);
                console.log(reply); // Log the reply instead of sending it
                // Uncomment the following line if you'd like the bot to send the response
                // await context.sendActivity(reply);
            } else {
                console.log("Type 'approved changes' to get recent approved changes.");
                // Uncomment the following line if you'd like the bot to send a default response
                // await context.sendActivity("Type 'approved changes' to get recent approved changes.");
            }
            await next();
        });

        // Send a welcome message
        this.onMembersAdded(async (context, next) => {
            await context.sendActivity("Hello! Type 'approved changes' to see approved changes in the past week.");
            await next();
        });
    }

    // Fetch approved changes from Freshservice
    async getApprovedChanges() {
        const lastWeekDate = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString();
        try {
            const response = await axios.get(`https://${process.env.FRESHSERVICE_DOMAIN}.freshservice.com/api/v2/changes`, {
                headers: {
                    'Authorization': `Basic ${Buffer.from(process.env.FRESHSERVICE_API_KEY + ':X').toString('base64')}`
                },
                params: {
                    status: 'scheduled',
                    updated_since: lastWeekDate
                },
                httpsAgent: httpsAgent  // Use the configured HTTPS agent for Zscaler
            });
            return response.data.changes;
        } catch (error) {
            console.error("Error fetching changes:", error);
            return [];
        }
    }

    // Format the response for Teams
    formatChangesForTeams(changes) {
        if (!changes.length) {
            return "No approved changes in the last 7 days.";
        }

        const cardContent = changes.map(change => ({
            title: change.title,
            subtitle: `Change Number: ${change.id}`,
            buttons: [
                {
                    type: 'openUrl',
                    title: 'View Change',
                    value: `https://${process.env.FRESHSERVICE_DOMAIN}.freshservice.com/change/${change.id}`
                }
            ]
        }));

        return MessageFactory.carousel(cardContent.map(change => CardFactory.heroCard(change.title, change.subtitle, null, change.buttons)));
    }
}

module.exports = FreshserviceBot;
