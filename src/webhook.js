const express = require("express");
const bodyParser = require("body-parser");
const WebSocket = require("ws");
const app = express();
const port = process.env.PORT || 3000;

app.use(express.json());
app.use(bodyParser.json());

let totalEvents = [];
const clients = [];

// Create HTTP server
const server = app.listen(port, () => {
   console.log(`Webhook server is running at http://localhost:${port}`);
});

// Create WebSocket server
const wss = new WebSocket.Server({ server });

// Handle WebSocket connections
wss.on("connection", (ws) => {
   console.log("Client connected via WebSocket");
   clients.push(ws);

   // Send the last 5 events from totalEvents array to the newly connected client
   ws.send(JSON.stringify(totalEvents.slice(-5)));

   ws.on("close", () => {
      console.log("Client disconnected");
      const index = clients.indexOf(ws);
      if (index !== -1) {
         clients.splice(index, 1);
      }
      totalEvents = []; // Clear the totalEvents array
   });
});

// Function to broadcast messages to all connected WebSocket clients
const broadcast = (message) => {
   clients.forEach((client) => {
      if (client.readyState === WebSocket.OPEN) {
         client.send(JSON.stringify(message));
      }
   });
};

// Function to process incoming webhook events asynchronously
const processEvents = async (events) => {
   await Promise.all(
      events.map(async (event) => {
         if (event.resourceData && event.resourceData.availability) {
            totalEvents.push(event.resourceData);
            const latestEvent = {
               status: event.resourceData.availability,
               activity: event.resourceData.activity,
               timestamp: new Date().toLocaleString("en-US"),
               id: event.resourceData.id,
            };
            console.log(
               "Updated status:",
               latestEvent.status,
               "and activity",
               latestEvent.activity,
               "at",
               latestEvent.timestamp,
               "for",
               latestEvent.id
            );
            broadcast(totalEvents.slice(-5)); // Broadcast the last 5 events from totalEvents array
         }
      })
   );
};

// Function to send heartbeat messages with totalEvents
const heartbeat = () => {
   if (clients.length > 0) {
      console.log("Sending heartbeat with latestEvents to all clients");
      broadcast(totalEvents.slice(-5)); // Broadcast the last 5 events from totalEvents array
   }
};

// Send heartbeat every 30 seconds
setInterval(heartbeat, 30000);

// Endpoint to receive webhook events
app.post("/notifications", (req, res) => {
   console.log("Received webhook event:", req.body);

   if (req.query.validationToken) {
      return res.status(200).send(req.query.validationToken);
   }

   const events = req.body.value || [];

   // Process events asynchronously to handle concurrency
   processEvents(events)
      .then(() => {
         res.status(200).send("Event received");
      })
      .catch((error) => {
         console.error("Error processing events:", error);
         res.status(500).send("Error processing events");
      });
});
