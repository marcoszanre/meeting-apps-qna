import * as debug from "debug";
import { v4 as uuidv4 } from 'uuid';

// tslint:disable-next-line:no-var-requires
const azure = require("azure-storage");

// Initialize debug logging module
const log = debug("msteams");

const tableSvc = azure.createTableService(process.env.STORAGE_ACCOUNT_NAME, process.env.STORAGE_ACCOUNT_ACCESSKEY);

const initTableSvc = () => {
    tableSvc.createTableIfNotExists("questionsTable", (error, result, response) => {
        if (!error) {
          // Table exists or created
          log("table service done");
        }
    });
};

const insertQuestion = (meetingid: string, author: string, question: string) => {

    return new Promise((resolve, reject) => { 
        const questionReference = {
            PartitionKey: {_: "questionsPartition"},
            RowKey: {_: uuidv4()},
            meetingid: {_: meetingid},
            author: {_: author},
            question: {_: question},
        };

        tableSvc.insertEntity("questionsTable", questionReference, (error, result, response) => {
            if (!error) {
              // Entity inserted
              log("success!");
              resolve("OK")
            } else {
              log(error);
              reject("Error")
            }
        });
    });
};

const deleteQuestion = (rowkey: string) => {

    return new Promise((resolve, reject) => { 
        const questionReference = {
            PartitionKey: {_: "questionsPartition"},
            RowKey: {_: rowkey}
        };

        tableSvc.deleteEntity("questionsTable", questionReference, (error, result, response) => {
            if (!error) {
              // Entity inserted
              log("success!");
              resolve("OK")
            } else {
              log(error);
              reject("Error")
            }
        });
    });
};


const getQuestions = async (meetingid: string, author: string) => {

    return new Promise((resolve) => {

        log("meeting id is " + meetingid);
        log("author is " + author);

        let myQuestions: Question[] = [];

        const query = new azure.TableQuery()
            .where("author eq ?", author)
            .and("meetingid eq ?", meetingid);

        tableSvc.queryEntities("questionsTable", query, null, (error, result) => {
            if (!error) {
              // query was successful
              for (let i = 0; i < result.entries.length; i++) {

                  const question: Question = {
                    meetingId: result.entries[i].meetingid._,
                    author: result.entries[i].author._,
                    question: result.entries[i].question._,
                    RowKey: result.entries[i].RowKey._
                  };

                  myQuestions.push(question);   
              }
            //   log(myQuestions.length);
            resolve(myQuestions);
            }
        });
    });
};

interface Question {
    meetingId: string;
    author: string;
    question: string;
    RowKey: string;
}


export {
    initTableSvc,
    insertQuestion,
    getQuestions,
    Question,
    deleteQuestion
}
