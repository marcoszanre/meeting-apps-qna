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
            promoted: {_: false},
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

const tableSvcUpdateQuestion = (rowkey: string, question: string) => {

    return new Promise((resolve, reject) => {
        const questionReference = {
            PartitionKey: {_: "questionsPartition"},
            RowKey: {_: rowkey},
            question: {_: question}
        };

        tableSvc.mergeEntity("questionsTable", questionReference, (error, result, response) => {
            if (!error) {
              // Entity inserted
              log("success!");
              resolve("OK");
            } else {
              log(error);
              reject("Error");
            }
        });
    });
};

const tableSvcPromoteDemoteQuestion = (rowkey: string, promoted: boolean) => {

    return new Promise((resolve, reject) => { 
        const questionReference = {
            PartitionKey: {_: "questionsPartition"},
            RowKey: {_: rowkey},
            promoted: {_: promoted}
        };

        tableSvc.mergeEntity("questionsTable", questionReference, (error, result, response) => {
            if (!error) {
              // Entity inserted
              log("success!");
              resolve("OK");
            } else {
              log(error);
              reject("Error");
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

const getAllQuestions = async (meetingid: string) => {

    return new Promise((resolve) => {

        // log("meeting id is " + meetingid);

        let myQuestions: Question[] = [];

        const query = new azure.TableQuery()
            .where("meetingid eq ?", meetingid);

        tableSvc.queryEntities("questionsTable", query, null, (error, result) => {
            if (!error) {
                // log(result);

              // query was successful
              for (let i = 0; i < result.entries.length; i++) {

                  const question: Question = {
                    meetingId: result.entries[i].meetingid._,
                    author: result.entries[i].author._,
                    question: result.entries[i].question._,
                    promoted: result.entries[i].promoted._,
                    RowKey: result.entries[i].RowKey._,
                    Timestamp: result.entries[i].Timestamp._
                  };

                  myQuestions.push(question);   
              }
            //   log(myQuestions.length);
            resolve(myQuestions);
            }
        });
    });
};

const setActiveQuestion = (meetingid: string, question: string) => {

    return new Promise(async (resolve, reject) => { 
        const questionReference = {
            PartitionKey: {_: "activeQuestionPartition"},
            RowKey: {_: meetingid},
            question: {_: question},
        };

        const isAlreadyActiveQuestion = await getActiveQuestion(meetingid);
        if (isAlreadyActiveQuestion === "no questions found") {
            // no active questions, create one
            tableSvc.insertEntity("questionsTable", questionReference, (error, result, response) => {
                if (!error) {
                  // Entity inserted
                  log("success!");
                  resolve("OK");
                } else {
                  log(error);
                  reject("Error");
                }
            });
        } else {
            // merge question reference as active question
            tableSvc.mergeEntity("questionsTable", questionReference, (error, result, response) => {
                if (!error) {
                  // Entity inserted
                  log("success!");
                  resolve("OK");
                } else {
                  log(error);
                  reject("Error");
                }
            });
        }

        
    });
};

const getActiveQuestion = async (rowkey: string) => {

    return new Promise((resolve, reject) => {

        tableSvc.retrieveEntity("questionsTable", "activeQuestionPartition", rowkey, (error, result, response) => {
            if (!error) {
                // result contains the entity
                const question = {
                    RowKey: result.RowKey._,
                    question: result.question._
                };
                // log(question);
                resolve(question.question);
            } else {
                resolve("no questions found");
            }
        });
    });
};

const setMeetingState = (rowkey: string, active: boolean) => {

    return new Promise(async (resolve, reject) => { 
        const meetingReference = {
            PartitionKey: {_: "meetingStatePartition"},
            RowKey: {_: rowkey},
            active: {_: active}
        };

        const isThereAlreadyAMeetingState = await getMeetingState(rowkey);
        if (isThereAlreadyAMeetingState === "not found") {
            // no active questions, create one
            tableSvc.insertEntity("questionsTable", meetingReference, (error, result, response) => {
                if (!error) {
                  // Entity inserted
                  log("success!");
                  resolve("OK");
                } else {
                  log(error);
                  reject("Error");
                }
            });
        } else {
            // merge question reference as active question
            tableSvc.mergeEntity("questionsTable", meetingReference, (error, result, response) => {
                if (!error) {
                  // Entity inserted
                  log("success!");
                  resolve("OK");
                } else {
                  log(error);
                  reject("Error");
                }
            });
        }

    });
};

const getMeetingState = async (rowkey: string) => {

    return new Promise((resolve, reject) => {

        tableSvc.retrieveEntity("questionsTable", "meetingStatePartition", rowkey, (error, result, response) => {
            if (!error) {
                // result contains the entity
                const meetingState = {
                    RowKey: result.RowKey._,
                    state: result.state._
                };
                // log(question);
                resolve(meetingState.state);
            } else {
                resolve("not found");
            }
        });
    });
};

interface Question {
    meetingId: string;
    author: string;
    question: string;
    RowKey: string;
    promoted?: boolean;
    Timestamp?: string;
}


export {
    initTableSvc,
    insertQuestion,
    getQuestions,
    Question,
    deleteQuestion,
    tableSvcUpdateQuestion,
    getAllQuestions,
    tableSvcPromoteDemoteQuestion,
    setActiveQuestion,
    getActiveQuestion,
    setMeetingState,
    getMeetingState
}
