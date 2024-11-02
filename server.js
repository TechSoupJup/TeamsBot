const restify = require('restify');
const { BotFrameworkAdapter } = require('botbuilder');
const FreshserviceBot = require('./bot');
require('dotenv').config();

// Create server
const server = restify.createServer();
server.listen(process.env.PORT || 3978, () => {
    console.log(`\nBot is listening on port ${process.env.PORT || 3978}`);
});

// Create adapter
const adapter = new BotFrameworkAdapter({
    appId: process.env.MICROSOFT_APP_ID,
    appPassword: process.env.MICROSOFT_APP_PASSWORD
});

// Error handling
adapter.onTurnError = async (context, error) => {
    console.error(`\n [onTurnError]: ${error}`);
    await context.sendActivity('Oops. Something went wrong!');
};

// Create the bot
const bot = new FreshserviceBot();

// Listen for incoming requests
server.post('/api/messages', async (req, res) => {
    await adapter.processActivity(req, res, async (context) => {
        await bot.run(context);
    });
});
