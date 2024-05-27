const express = require("express");
const bodyParser = require("body-parser");
const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());
app.use(bodyParser.json());

let totalEvents = [];

// Endpoint to receive webhook events
app.post("/notifications", (req, res) => {
   console.log("Received webhook event:", req.body);

   if (req.query.validationToken) {
      return res.status(200).send(req.query.validationToken);
   }

   const events = req.body.value || [];
   events.forEach((event) => {
      if (event.resourceData && event.resourceData.availability) {
         totalEvents.push(event.resourceData);
         const latestEvent = {
            status: event.resourceData.availability,
            timestamp: new Date().toLocaleString("en-US"),
            id: event.resourceData.id,
         };
         console.log(
            "Updated status:",
            latestEvent.status,
            "at",
            latestEvent.timestamp,
            "for",
            latestEvent.id
         );
      }
   });
   res.status(200).send("Event received");
});

app.get("/events", (req, res) => {
   res.status(200).json(totalEvents);
});

app.listen(port, () => {
   console.log(`Webhook server is running at http://localhost:${port}`);
});
