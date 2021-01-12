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

    tableSvc.createTableIfNotExists("likesTable", (error, result, response) => {
        if (!error) {
          // Table exists or created
          log("table service likes done");
        }
    });
    
    tableSvc.createTableIfNotExists("auditTable", (error, result, response) => {
        if (!error) {
          // Table exists or created
          log("table service audit done");
        }
    });

    tableSvc.createTableIfNotExists("meetingsTable", (error, result, response) => {
        if (!error) {
          // Table exists or created
          log("table service audit done");
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
            likedBy: {_: 0},
            asked: {_: false},
            askedWhen: {_: "not asked yet"}
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
                    RowKey: result.entries[i].RowKey._,
                    likedBy: result.entries[i].likedBy._
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
                    Timestamp: result.entries[i].Timestamp._,
                    likedBy: result.entries[i].likedBy._,
                    asked: result.entries[i].asked._,
                    askedWhen: result.entries[i].askedWhen._
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
            // no active meeting state, create one
            tableSvc.insertEntity("meetingsTable", meetingReference, (error, result, response) => {
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
            tableSvc.mergeEntity("meetingsTable", meetingReference, (error, result, response) => {
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

        tableSvc.retrieveEntity("meetingsTable", "meetingStatePartition", rowkey, (error, result, response) => {
            if (!error) {
                // result contains the entity
                const meetingState = {
                    RowKey: result.RowKey._,
                    active: result.active._
                };
                // log(question);
                resolve(meetingState.active);
            } else {
                resolve("not found");
            }
        });
    });
};

const getLike = async (questionId: string, userID: string) => {

    log("get liked called");
    log(questionId);
    log(userID);


    return new Promise((resolve) => {

        const query = new azure.TableQuery()
            .where("userID eq ?", userID)
            .and("questionId eq ?", questionId);

        tableSvc.queryEntities("likesTable", query, null, (error, result) => {
            log("the result is:");
            log(result);

            if (!error) {
              // query was successful
              log(result.entries.length);
              if (result.entries.length == 0) {
                log("like NOT found");
                resolve(false);
              } else {
                log("like found");
                resolve(true);
              }
            } else {
              log("like NOT found");
              resolve(false);
            }
        });
    });
};

const getLikeRow = async (questionId: string, userID: string) => {

    log("get liked row called");
    log(questionId);
    log(userID);


    return new Promise((resolve) => {

        const query = new azure.TableQuery()
            .where("userID eq ?", userID)
            .and("questionId eq ?", questionId);

        tableSvc.queryEntities("likesTable", query, null, (error, result) => {
            log("the result is:");
            //log(result);

            if (!error) {
              // query was successful
              resolve(result.entries[0].RowKey._)
            } else {
              resolve(false);
            }
        });
    });
};

const insertLike = async (questionId: string, userID: string) => {

    return new Promise((resolve) => {

            const likeReference = {
                PartitionKey: {_: "likesPartition"},
                RowKey: {_: uuidv4()},
                questionId: {_: questionId},
                userID: {_: userID}
            };

            tableSvc.insertEntity("likesTable", likeReference, (error, result, response) => {
                if (!error) {
                  // Entity inserted
                  log("success!");
                  resolve("OK");
                } else {
                  log(error);
                  resolve("Error");
                }
            });
    });
};

const removeLike = async (questionId: string, userID: string) => {

    log("remove like called");

    return new Promise(async (resolve) => {

        const rowkey = await getLikeRow(questionId, userID);
        // log(rowkey);

        const likeReference = {
            PartitionKey: {_: "likesPartition"},
            RowKey: {_: rowkey}
        };

        tableSvc.deleteEntity("likesTable", likeReference, (error, result, response) => {
            if (!error) {
              // Entity inserted
              log("success!");
              resolve("OK")
            } else {
              log(error);
              resolve("Error")
            }
        });

    });
};

const updateLikeAggregate = async (questionId: string) => { 

    let likeCount: number;

    const query = new azure.TableQuery()
            .where("questionId eq ?", questionId)

    tableSvc.queryEntities("likesTable", query, null, (error, result) => {
            // log("the result is:");
            // log(result);

            if (!error) {
              // query was successful
              likeCount = result.entries.length;

              // update table like count
              const questionReference = {
                PartitionKey: {_: "questionsPartition"},
                RowKey: {_: questionId},
                likedBy: {_: likeCount}
              };

            tableSvc.mergeEntity("questionsTable", questionReference, (error, result, response) => {
                if (!error) {
                  // Entity inserted
                  log("success!");
                } else {
                  log(error);
                }
            });
        }
    });

}

const toggleLike = async (questionId: string, userID: string) => {

        const isLiked = await getLike(questionId, userID);
        log("isLiked = " + isLiked);

        if (isLiked) {
            await removeLike(questionId, userID);

        } else {
            await insertLike(questionId, userID);
        }

        await updateLikeAggregate(questionId);
};

const tableSvcSetAskedQuestion = (rowkey: string) => {

    return new Promise((resolve, reject) => { 
        const questionReference = {
            PartitionKey: {_: "questionsPartition"},
            RowKey: {_: rowkey},
            asked: {_: true},
            askedWhen: {_: Date.now().toLocaleString()}
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

interface Question {
    meetingId: string;
    author: string;
    question: string;
    RowKey: string;
    promoted?: boolean;
    Timestamp?: string;
    likedBy?: number;
    asked?: boolean;
    askedWhen?: string;
}


const auditAction = async (action: string, actor: string, content: string, meetingid: string) => {

    const auditReference = {
        PartitionKey: {_: "auditPartition"},
        RowKey: {_: uuidv4()},
        action: {_: action},
        actor: {_: actor},
        content: {_: content},
        meetingid: {_: meetingid}
    };

    tableSvc.insertEntity("auditTable", auditReference, (error, result, response) => {
        if (!error) {
          // Entity inserted
          log("success!");
        } else {
          log(error);
        }
    });

};


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
    getMeetingState,
    getLike,
    toggleLike,
    tableSvcSetAskedQuestion
}
