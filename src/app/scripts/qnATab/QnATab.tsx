import * as React from "react";
import { Provider, Flex, Button, Header, Segment, Alert, TextArea, List, ButtonGroup, Loader, ListProps, Dialog, Input, Table } from "@fluentui/react-northstar";
import { CloseIcon, EditIcon } from '@fluentui/react-icons-northstar';
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwt_decode from "jwt-decode";
import { Question } from "../../services/tableService";

/**
 * Implementation of the QnA content page
 */
export const QnATab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    const [question, setQuestion] = useState<string>();
    const [myQuestions, setMyQuestions] = useState<listItem[]>();
    const [btnLoading, setBtnLoading] = useState<boolean>(false);
    const [selectedListIndex, setSelectedListIndex] = useState<number>(-1);
    const [removeBtnDisabled, setRemoveBtnDisabled] = useState<boolean>(true);
    const [isRemoveDialogOpen, setIsRemoveDialogOpen] = useState<boolean>(false);
    const [dialogContent, setDialogContent] = useState<string>();


    useEffect(() => {
        if (inTeams === true) {

            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    console.log(token);
                    const decoded: { [key: string]: any; } = jwt_decode(token) as { [key: string]: any; };
                    setName(decoded!.name);
                    microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [process.env.QNA_APP_URI as string]
            });
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
            console.log(context);
        }
    }, [context]);

    // call load questions when once name is set
    useEffect(() => {
        if (name && context) {
            updateQuestions();
        }
    }, [name])

    const updateQuestions = async () => {

        let listItems: listItem[] = [];
            loadQuestions().then((result: Question[]) => {
                for (let index = 0; index < result.length; index++) {
                    
                    const listItem: listItem = {
                        content: result[index].question,
                        key: result[index].RowKey
                        // endMedia: actions
                    }

                    listItems.push(listItem);
                }

                setMyQuestions(listItems);
            });
            // console.log("ready to go " + name + "///////////" +context.meetingId);
    }

    const remoteQuestionFromArray = async () => {
        setMyQuestions(myQuestions?.filter((item) => item.key !== myQuestions![selectedListIndex!].key));
    }

    const loadQuestions = async () => {

        // console.log("Your meeting id is" + context?.meetingId);
        // console.log("Your name is " + name);
        // console.log(fetchUrl);

        const myMeetingId: string = context?.meetingId as string;
        const myName: string = name as string;
        const fetchUrl: string = `/api/question?meetingId=${myMeetingId}&author=${myName}`;
        const questionsList = await fetch(fetchUrl);        
        return questionsList.json();
    };

    const actions = (
        <ButtonGroup
          buttons={[
            {
              icon: <EditIcon />,
              iconOnly: true,
              text: true,
              key: "check",
              onClick: (() => {
                console.log("edit " + selectedListIndex);
              })
            },
            {
              icon: <CloseIcon />,
              iconOnly: true,
              text: true,
              key: "close",
              onClick: (() => {
                console.log("close " + selectedListIndex);
              })
            },
          ]}
        />
      );

    interface listItem {
        key: string;
        // endMedia: any;
        content: string;
    }

    // const listItems: listItem[] = [
    //     {
    //       key: "sensor",
    //       endMedia: actions,
    //       content: "Program the sensor to the SAS alarm through the haptic SQL card!"
    //     },
    //     {
    //       key: "ftp",
    //       endMedia: actions,
    //       content: "Use the online FTP application to input the multi-byte application!"
    //     },
    //     {
    //       key: "gb",
    //       endMedia: actions,
    //       content: "The GB pixel is down, navigate the virtual interface!"
    //     }
    // ];

    const listItems: listItem[] = myQuestions as listItem[];

    const ListExample = () => <List 
                                selectable
                                selectedIndex={selectedListIndex}
                                onSelectedIndexChange={(e, newProps: ListProps) => {
                                setSelectedListIndex(newProps.selectedIndex as number);
                                setRemoveBtnDisabled(false);
                                // console.log(selectedListIndex);
                                }}
                                styles={{
                                    paddingTop: "0.625rem"
                                }}
                                items={listItems} variables={{
                                    rootPadding: "0rem"
    }}/>

    const submitQuestion = async () => {
        // console.log(question);
        setBtnLoading(true);
        await postQuestion();
        setQuestion("");
        updateQuestions();
        setBtnLoading(false);
    };

    const handleChange = (event) => {
        setQuestion(event.target.value);
    };

    // const handleEditBtnClicked = () => {
    //     setIsEditDialogOpen(false);
    //     console.log("edit list item " + myQuestions![selectedListIndex!].content);
    // };

    const handleRemoveBtnClicked = async () => {
        console.log("remove button clicked " + dialogContent);
        setIsRemoveDialogOpen(false);

        const res = await fetch(`/api/question?rowkey=${myQuestions![selectedListIndex!].key}`, {
            method: "DELETE"
        });

        console.log(res.status);
        if (res.status === 200) {
            remoteQuestionFromArray();
        } else (console.log("an error happened"));

    };


    const postQuestion = async () => {
        const questionData = {
            meetingId: context?.meetingId,
            author: name,
            question: question
        };

        const body = JSON.stringify(questionData);
        // console.log(body);

        const res = await fetch("/api/question", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: body
        });

        // console.log(res.status);
    };

    const header = {
        key: 'header',
        items: [
          { content: 'id', key: 'id' },
          { content: 'Author', key: 'name' },
          { content: 'Question', key: 'pic' },
          { content: 'Selected', key: 'action' },
        ],
      };
      const rowsPlain = [
        {
          key: 1,
          items: [
            { content: '1', key: '1-1' },
            {
              content: 'Roman van von der Longername',
              truncateContent: true,
              key: '1-2',
            },
            { content: 'None', key: '1-3' },
            { content: 'yes', key: '1-4' },
          ],
        },
        {
          key: 2,
          items: [
            { content: '2', key: '2-1' },
            { content: 'Alex', key: '2-2' },
            { content: 'None', key: '2-3' },
            { content: 'yes', key: '2-4' },
          ],
        },
        {
          key: 3,
          items: [
            { content: '3', key: '3-1' },
            { content: 'Ali', key: '3-2' },
            { content: 'None', key: '3-3' },
            { content: 'no', truncateContent: true, key: '3-4' },
          ],
        },
      ];

    /**
     * The render() method to create the UI of the tab
     */
    return (
        <Provider theme={theme}>
            
            <Flex column padding="padding.medium">

                <Header
                    as="h2"
                    content="Meeting QnA"
                    description={{
                    content: "Use the options below to create or manage your questions",
                    as: "span",
                    }}
                    styles={{
                        paddingLeft: "1.250rem",
                        paddingBottom: "0.625rem"
                }}/>
                
                <Segment>
                        <Header as="h2" content="All Questions" />
                        <Flex gap="gap.small">
                            <Input fluid placeholder="Search..." />
                        </Flex>
                        <Table variables={{ cellContentOverflow: 'none' }} header={header} rows={rowsPlain} aria-label="Static table" />
                    </Segment>
            </Flex>

            <Flex gap="gap.small" padding="padding.medium">
                <Flex.Item size="size.half">
                    <Segment>
                        <Header as="h2" content="My Questions" />
                        <Flex gap="gap.small">
                            {/* <Button disabled={editAndRemoveBtnDisabled} onClick={handleEditBtnClicked} icon={<EditIcon />} content="Old Edit" iconPosition="before" /> */}
                            {/* <Button disabled={editAndRemoveBtnDisabled} onClick={handleRemoveBtnClicked} icon={<CloseIcon />} content="Remove" iconPosition="before" /> */}
                            <Dialog
                                open={isRemoveDialogOpen}
                                onOpen={() => { setIsRemoveDialogOpen(true); setDialogContent(myQuestions![selectedListIndex!].content) }}
                                onCancel={() => setIsRemoveDialogOpen(false)}
                                onConfirm={handleRemoveBtnClicked}
                                confirmButton="Confirm"
                                cancelButton="Cancel"
                                content={dialogContent}
                                header="Remove this question?"
                                headerAction={{
                                    icon: <CloseIcon />,
                                    title: 'Close',
                                    onClick: () => setIsRemoveDialogOpen(false),
                            }}
                            trigger={<Button disabled={removeBtnDisabled} icon={<CloseIcon />} content="Remove" iconPosition="before" />}
                            />
                        </Flex>
                        {myQuestions ? <ListExample /> : <Loader label="Loading questions" />}
                    </Segment>
                </Flex.Item>
                <Flex.Item size="size.half">
                    <Segment>
                        <Header as="h2" content="Ask a Question" />
                        <TextArea disabled={btnLoading} value={question} onChange={handleChange} fluid placeholder="Type your question here..." resize="vertical"/>
                        <Button loading={btnLoading} onClick={submitQuestion} content="Submit" primary styles={{
                            float: "right",
                            marginTop: "0.625rem"
                        }}/>
                    </Segment>
                </Flex.Item>
            </Flex>

        </Provider>
    );
};
