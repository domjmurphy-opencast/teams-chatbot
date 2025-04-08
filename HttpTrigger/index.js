const axios = require('axios');
const { Configuration, OpenAIApi } = require("openai");

module.exports = async function (context, req) {
  context.log('HTTP trigger function processed a request.');

  const query = req.body && req.body.query;
  if (!query) {
    context.res = {
      status: 400,
      body: "Please pass a 'query' in the request body"
    };
    return;
  }

  let handbookText = "";

  // Check if SharePoint integration is disabled
  const disableSharepoint = process.env.DISABLE_SHAREPOINT === 'true';

  // Attempt SharePoint retrieval only if enabled and configured
  if (!disableSharepoint && process.env.SHAREPOINT_SITE_ID && process.env.SHAREPOINT_DOC_ID && process.env.GRAPH_ACCESS_TOKEN) {
    try {
      const spResponse = await axios.get(
        `https://graph.microsoft.com/v1.0/sites/${process.env.SHAREPOINT_SITE_ID}/drive/items/${process.env.SHAREPOINT_DOC_ID}/content`,
        {
          headers: {
            Authorization: `Bearer ${process.env.GRAPH_ACCESS_TOKEN}`
          },
          responseType: 'text'
        }
      );
      handbookText = spResponse.data;
    } catch (error) {
      context.log.error("Error fetching handbook from SharePoint:", error.message);
      handbookText = "Could not retrieve handbook from SharePoint. Please ensure the document is available.";
    }
  } else {
    context.log("SharePoint integration is disabled or not configured. Using fallback handbook text.");
    handbookText = "Fallback handbook content: Insert your handbook content here for testing purposes.";
  }

  // --- Step 2: Call OpenAI Chat Completion API ---
  const configuration = new Configuration({
    apiKey: process.env.OPENAI_API_KEY,
  });
  const openai = new OpenAIApi(configuration);

  // Build messages for the chat completion API
  const messages = [
    { role: "system", content: `You are an expert on the following employee handbook:\n\n${handbookText}` },
    { role: "user", content: `Question: ${query}` }
  ];

  try {
    const completion = await openai.createChatCompletion({
      model: "gpt-3.5-turbo",
      messages: messages,
      max_tokens: 150,
      temperature: 0.2,
    });

    const answer = completion.data.choices[0].message.content.trim();
    context.res = {
      body: { answer }
    };
  } catch (error) {
    context.log.error("Error calling OpenAI API:", error.message);
    context.res = {
      status: 500,
      body: "Error generating answer."
    };
  }
};
