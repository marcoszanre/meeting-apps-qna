import { Flex, Header, Segment, Input, Dialog, CloseIcon, Button, Loader, TextArea, List, ListProps, TrashCanIcon, EditIcon, RetryIcon } from "@fluentui/react-northstar";
import { Context } from "@microsoft/teams-js";
import * as React from "react";
import { useState, useEffect } from "react";
import { Question } from "../../../services/tableService";

interface IOrganizerProps {
    context: Context;
    name: string;
}

export const Atendee: React.FC<IOrganizerProps> = ({ context, name }) => {

    const [question, setQuestion] = useState<string>();
    const [myQuestions, setMyQuestions] = useState<IListItem[]>();
    const [btnLoading, setBtnLoading] = useState<boolean>(false);
    const [btnUpdateLoading, setbtnUpdateLoading] = useState<boolean>(false);
    const [selectedListIndex, setSelectedListIndex] = useState<number>(-1);
    const [removeBtnDisabled, setRemoveBtnDisabled] = useState<boolean>(true);
    const [isRemoveDialogOpen, setIsRemoveDialogOpen] = useState<boolean>(false);
    const [isNotEditingQuestion, setIsNotEditingQuestion] = useState<boolean>(true);
    const [dialogContent, setDialogContent] = useState<string>();
    const [editedQuestion, setEditedQuestion] = useState<string>();
    const [isNotAllowedToSubmitQuestion, setIsNotAllowedToSubmitQuestion] = useState<boolean>(true);



    interface IListItem {
        key: string;
        content: string;
    }

    useEffect(() => {
        loadQuestions();
        updateQuestions();
    }, []);

    const listItems: IListItem[] = myQuestions as IListItem[];

    const updateQuestions = async () => {

        const listItems: IListItem[] = [];
        loadQuestions().then((result: Question[]) => {
            for (let index = 0; index < result.length; index++) {

                const listItem: IListItem = {
                    content: result[index].question,
                    key: result[index].RowKey
                };

                listItems.push(listItem);
            }

            setMyQuestions(listItems);
        });
    };

    const remoteQuestionFromArray = async () => {
        setMyQuestions(myQuestions?.filter((item) => item.key !== myQuestions![selectedListIndex!].key));
    };

    const loadQuestions = async () => {


        const myMeetingId: string = context?.meetingId as string;
        const myName: string = name as string;
        const fetchUrl: string = `/api/question?meetingId=${myMeetingId}&author=${myName}`;
        const questionsList = await fetch(fetchUrl);
        return questionsList.json();
    };

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
    }}/>;

    const submitQuestion = async () => {
        // console.log(question);
        setBtnLoading(true);
        await postQuestion();
        setQuestion("");
        updateQuestions();
        setBtnLoading(false);
        setIsNotAllowedToSubmitQuestion(true);
    };

    const updateQuestion = async () => {
        // console.log(question);
        setbtnUpdateLoading(true);
        await updateTableQuestion();
        setEditedQuestion("");
        await updateQuestions();
        setbtnUpdateLoading(false);
    };

    const handleChange = (event) => {
        setQuestion(event.target.value);
        if (question?.length && question.length > 3) {
            setIsNotAllowedToSubmitQuestion(false);
        } else if (typeof(question?.length) === undefined || question?.length! < 3) {
            setIsNotAllowedToSubmitQuestion(true);
        }
    };

    const handleEditChange = (event) => {
        setEditedQuestion(event.target.value);
    };

    const handleEditBtnClicked = () => {
        setEditedQuestion(myQuestions![selectedListIndex!].content);
        setIsNotEditingQuestion(false);
        // console.log("edit clicked");
        // console.log("edit btn clicked " + myQuestions![selectedListIndex!].content);
    };

    const handleRemoveBtnClicked = async () => {
        // console.log("remove button clicked " + dialogContent);
        setIsRemoveDialogOpen(false);

        const res = await fetch(`/api/question?rowkey=${myQuestions![selectedListIndex!].key}`, {
            method: "DELETE"
        });

        // console.log(res.status);
        if (res.status === 200) {
            remoteQuestionFromArray();
        } else {
            // console.log("an error happened");
        }

    };

    const handleCancelBtnClicked = async () => {
        // console.log("cancel button clicked");
        !isNotEditingQuestion && setIsNotEditingQuestion(true);
        setSelectedListIndex(-1);
        setEditedQuestion("");
        setRemoveBtnDisabled(true);
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

    const updateTableQuestion = async () => {

        const rowkey: string = myQuestions![selectedListIndex!].key as string;
        const fetchUrl: string = `/api/question?rowkey=${rowkey}&question=${editedQuestion}`;

        const questionUpdateReponse = await fetch(fetchUrl, {
            method: "PATCH"
        });

        if (questionUpdateReponse.status === 200) {
            setEditedQuestion("");
            setIsNotEditingQuestion(true);
        } else {
            // console.log("an error happened");
        }

        // console.log(questionUpdateReponse);

    };

    return (
        <>
        <Flex column padding="padding.medium">

        <Header
            as="h2"
            content="Meeting QnA - Atendee"
            description={{
            content: "Use the options below to create or manage your questions",
            as: "span",
            }}
            styles={{
                paddingLeft: "1.250rem"
                // paddingBottom: "0.625rem"
        }}/>
        <RetryIcon title="Refresh Questions" onClick={updateQuestions} styles={{
                    position: "absolute",
                    right: "0",
                    marginTop: "1.625rem",
                    marginRight: "1.625rem",
                    top: "0",
                    cursor: "pointer"
        }}/>

        <Segment>
                <Header as="h2" content="Ask a question" />
                <TextArea disabled={btnLoading} value={question} onChange={handleChange} fluid placeholder="Type your question here..." resize="vertical"/>
                <Button disabled={isNotAllowedToSubmitQuestion} loading={btnLoading} onClick={submitQuestion} content="Submit" primary styles={{
                    float: "right",
                    marginTop: "0.625rem"
                }}/>
            </Segment>
    </Flex>

    <Flex gap="gap.small" padding="padding.medium">
        <Flex.Item size="size.half">
            <Segment>
                <Header as="h2" content="My questions" />
                <Flex gap="gap.small">
                    <Button disabled={removeBtnDisabled} onClick={handleEditBtnClicked} icon={<EditIcon />} content="Edit" iconPosition="before" />
                    <Dialog
                        open={isRemoveDialogOpen}
                        onOpen={() => { setIsRemoveDialogOpen(true); setDialogContent(myQuestions![selectedListIndex!].content); }}
                        onCancel={() => setIsRemoveDialogOpen(false)}
                        onConfirm={handleRemoveBtnClicked}
                        confirmButton="Confirm"
                        cancelButton="Cancel"
                        content={dialogContent}
                        header="Remove this question?"
                        headerAction={{
                            icon: <CloseIcon />,
                            title: "Close",
                            onClick: () => setIsRemoveDialogOpen(false),
                    }}
                    trigger={<Button disabled={removeBtnDisabled} icon={<TrashCanIcon />} content="Remove" iconPosition="before" />}
                    />
                    <Button disabled={removeBtnDisabled} onClick={handleCancelBtnClicked} icon={<CloseIcon />} content="Cancel" iconPosition="before" />
                </Flex>
                {myQuestions ? <ListExample /> : <Loader label="Loading questions" />}
            </Segment>
        </Flex.Item>
        <Flex.Item size="size.half">
            <Segment>
                <Header as="h2" content="Edit question" />
                <TextArea disabled={isNotEditingQuestion} value={editedQuestion} onChange={handleEditChange} fluid placeholder="Select a question to edit it." resize="vertical"/>
                <Button disabled={isNotEditingQuestion} loading={btnUpdateLoading} onClick={updateQuestion} content="Save" primary styles={{
                    float: "right",
                    marginTop: "0.625rem"
                }}/>
            </Segment>
        </Flex.Item>
    </Flex>
    </>
  );
};
