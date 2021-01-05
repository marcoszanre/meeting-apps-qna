import * as Express from "express";
import * as http from "http";
import * as path from "path";
import * as morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import * as debug from "debug";
import * as compression from "compression";
import { initTableSvc, insertQuestion, getQuestions, deleteQuestion, tableSvcUpdateQuestion, getAllQuestions, tableSvcPromoteDemoteQuestion, setActiveQuestion, getActiveQuestion, setMeetingState, getMeetingState } from "./services/tableService";



// Initialize debug logging module
const log = debug("msteams");

log(`Initializing Microsoft Teams Express hosted App...`);

// Initialize dotenv, to use .env file settings if existing
// tslint:disable-next-line:no-var-requires
require("dotenv").config();



// The import of components has to be done AFTER the dotenv config
import * as allComponents from "./TeamsAppsComponents";

// Create the Express webserver
const express = Express();
const port = process.env.port || process.env.PORT || 3007;

// Inject the raw request body onto the request object
express.use(Express.json({
    verify: (req, res, buf: Buffer, encoding: string): void => {
        (req as any).rawBody = buf.toString();
    }
}));
express.use(Express.urlencoded({ extended: true }));

// Express configuration
express.set("views", path.join(__dirname, "/"));

// Add simple logging
express.use(morgan("tiny"));

// Add compression - uncomment to remove compression
express.use(compression());

// Add /scripts and /assets as static folders
express.use("/scripts", Express.static(path.join(__dirname, "web/scripts")));
express.use("/assets", Express.static(path.join(__dirname, "web/assets")));

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsApiRouter(allComponents));

express.use("/api/question", async (req, res) => {
    if (req.method === "POST") {

        // log(`Contoso Post Request`);
        // log(req);

        const meetingId = req.body.meetingId;
        const author = req.body.author;
        const question = req.body.question;

        // log(req.body);
        // log(meetingId);
        // log(author);
        // log(question);

        const response  = await insertQuestion(meetingId, author, question);

        log(response);
        response === "OK" ? res.status(201) : res.status(500);
        res.send();
    } else if (req.method === "GET") {

        // handle get request
        // log("get request called");
        // log(req.query);

        const meetingId = req.query.meetingId as string;
        const author = req.query.author as string;

        // log(meetingId);
        // log(author);

        let questionsList;

        if (author === "all") {
            questionsList = await getAllQuestions(meetingId);
        } else {
            questionsList = await getQuestions(meetingId, author);
        }

        // log(questionsList);
        res.json(questionsList);
    } else if (req.method === "DELETE") {

        // handle DELETE request
        log("delete request called");
        // log(req.query);

        const rowkey = req.query.rowkey as string;

        const response = await deleteQuestion(rowkey);
        log(response);
        response === "OK" ? res.status(200) : res.status(500);
        res.send();
    } else if (req.method === "PATCH") {

        // handle DELETE request
        log("PATCH request called");
        // log(req.query);

        const rowkey = req.query.rowkey as string;
        let response;

        if (req.query.question) {
            // handle update question event
            const question = req.query.question as string;
            response = await tableSvcUpdateQuestion(rowkey, question);

        } else {
            // handle promote question
            const promoted: boolean = (req.query.promoted === "true");
            response = await tableSvcPromoteDemoteQuestion(rowkey, promoted);
        }

        log(response);
        response === "OK" ? res.status(200) : res.status(500);
        res.send();
    }
});

express.use("/api/bubble", async(req, res, next) => {
    
    if (req.method === "POST") {

        log("POST bubble called");
        
        const chatId = req.body.chatId;
        const author = req.body.author;
        const question = req.body.question;

        const meetingBubbleTitle = `Question from ${author}`;
        
        const accessToken = await getAuthTokenFromMicrosoft();
        await sendBubbleMessage(accessToken, chatId, meetingBubbleTitle, question, author);

        res.status(200)
        res.send();
    }

    // next();
});

express.use("/api/activequestion", async(req, res, next) => {
    
    if (req.method === "PATCH") {

        log("PATCH active question called");
        
        const meetingid = req.body.meetingid;
        const question = req.body.question;

        const response = await setActiveQuestion(meetingid, question);        
        // log(response);

        res.status(200)
        res.send();

    } else if ( req.method === "GET") {

        log("GET active question called");

        const meetingid = req.query.meetingid;
        // log(meetingid);

        const activeQuestionContent = await getActiveQuestion(meetingid as string);
        log(activeQuestionContent);
        res.json({ activeQuestion: activeQuestionContent });

    }

    // next();
});

express.use("/api/meetingstate", async(req, res, next) => {
    
    if (req.method === "POST") {

        log("POST meeting state called");

        const meetingid = req.body.meetingid;
        const active = req.body.active;

        const response = await setMeetingState(meetingid, active);
        // log(response);

        res.status(200)
        res.send();

    } else if (req.method === "GET") {

        log("GET meeting state called");

        const meetingid = req.query.meetingid;
        // log(meetingid);

        const meetingState = await getMeetingState(meetingid as string);
        log(meetingState);
        res.json({ meetingState: meetingState });

    }

});


// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
express.use(MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents
}));

// Set default web page
express.use("/", Express.static(path.join(__dirname, "web/"), {
    index: "index.html"
}));

// Set the port
express.set("port", port);

// Start the webserver
http.createServer(express).listen(port, () => {
    log(`Server running on ${port}`);
});

// Init table service
initTableSvc();


// Send get user role
express.get("/api/role", async (req, res, next) => {

    // console.log(req.query);

    const meetingId = req.query.meetingId;
    const userId = req.query.userId;

    // const accessToken = await getAuthTokenFromMicrosoft(process.env.MICROSOFT_APP_ID, process.env.MICROSOFT_APP_PASSWORD);
    const accessToken = await getAuthTokenFromMicrosoft();
    // log(accessToken);
    const memberRoleResponse = await getMeetingParticipant(accessToken, meetingId, userId, process.env.TENANT_ID as string);
    // log(memberRoleResponse);
    const userRole = memberRoleResponse.meetingRole;

    // console.log(meetingId);
    // console.log(userId);

    await res.json({ role: userRole });
    next();
});

// Sends the bubble notification
// express.get("/api/bubble", async (req, res, next) => {

//     // console.log(req.query);

//     const chatId = req.query.chatId;
//     const meetingBubbleTitle = "Contoso";

//     // const accessToken = await getAuthTokenFromMicrosoft(process.env.MICROSOFT_APP_ID, process.env.MICROSOFT_APP_PASSWORD);
//     const accessToken = await getAuthTokenFromMicrosoft();
//     await sendBubbleMessage(accessToken, chatId, meetingBubbleTitle);

//     res.sendStatus(200);
//     next();
// });


// const getAuthTokenFromMicrosoft = async (appId, appSecret) => {
const getAuthTokenFromMicrosoft = async () => {

    const details = {
        scope: "https://api.botframework.com/.default",
        grant_type: "client_credentials",
        client_id: process.env.MICROSOFT_APP_ID,
        client_secret: process.env.MICROSOFT_APP_PASSWORD
    };

    const formBody: string[] = [];

    for (const property in details) {
      if (details.hasOwnProperty(property)) {
        const encodedKey = encodeURIComponent(property);
        const encodedValue = encodeURIComponent(details[property]);
        formBody.push(encodedKey + "=" + encodedValue);
      }
    }

    const postBody = formBody.join("&");

    const res = await fetch("https://login.microsoftonline.com/botframework.com/oauth2/v2.0/token", {
      method: "POST",
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
      },
      body: postBody,
    });
    const json = await res.json();
    if (json.error) { throw new Error(`${json.error}: ${json.error_description}`); }
    // log(json);
    return json.access_token;
};

const getMeetingParticipant = async (token, meetingId, participantId, tenandId) => {

    // /v1/meetings/{meetingId}/participants/{participantId}?tenantId={tenantId}
    const res = await fetch(`https://smba.trafficmanager.net/amer/v1/meetings/${meetingId}/participants/${participantId}?tenantId=${tenandId}`, {
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      }
    });
    const json = await res.json();
    if (json.error) { throw new Error(`${json.error}: ${json.error_description}`); }
    return json;

};

const sendBubbleMessage = async (token, chatid, meetingBubbleTitle, question, author) => {

    // /v1/meetings/{meetingId}/participants/{participantId}?tenantId={tenantId}

    const meetingTabURL = `https://${process.env.HOSTNAME}/qnATab/?name={loginHint}&tenant={tid}&group={groupId}&theme={theme}`;
    const res = await fetch(`https://smba.trafficmanager.net/amer/v3/conversations/${chatid}/activities`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`
      },
      body: JSON.stringify({
        type: "message",
        attachments: [
            {
                "contentType": "application/vnd.microsoft.card.hero",
                "content": {
                    "title": "QnA",
                    "subtitle": author,
                    "text": `"${question}"`,
                }
            }
        ],
        channelData: {
            notification: {
                alertInMeeting: true,
                externalResourceUrl: `https://teams.microsoft.com/l/bubble/${process.env.APPLICATION_ID}?url=${meetingTabURL}/&height=300&width=428&title=${meetingBubbleTitle}&completionBotId=${process.env.MICROSOFT_APP_ID}`
            }
        }
    })
    });
    const json = await res.json();
    if (json.error) { throw new Error(`${json.error}: ${json.error_description}`); }
    return json;

};
