import { Flex, Header, Segment, Input, Dialog, CloseIcon, Button, Loader, TextArea, List, ListProps, TrashCanIcon, EditIcon, AcceptIcon, EyeIcon, EyeSlashIcon, SearchIcon, RetryIcon } from "@fluentui/react-northstar";
import { Context } from "@microsoft/teams-js";
import * as React from "react";
import { useState, useEffect } from "react";
import { Question } from "../../../services/tableService";

type OrganizerProps = {
    context: Context,
    name: string,
}

export const Organizer: React.FC<OrganizerProps> = ({ context, name }) => {

    const [question, setQuestion] = useState<string>();
    const [allQuestions, setMyQuestions] = useState<listItem[]>();
    const [promotedQuestions, setPromotedQuestions] = useState<listItem[]>();
    const [notPromotedQuestions, setNotPromotedQuestions] = useState<listItem[]>();
    const [btnLoading, setBtnLoading] = useState<boolean>(false);
    const [btnUpdateLoading, setbtnUpdateLoading] = useState<boolean>(false);
    const [selectedPromotedListIndex, setSelectedPromotedListIndex] = useState<number>(-1);
    const [selectedNotPromotedListIndex, setSelectedNotPromotedListIndex] = useState<number>(-1);
    const [removeBtnDisabled, setRemoveBtnDisabled] = useState<boolean>(true);
    const [notPromotedBtnsDisabled, setNotPromotedBtnsDisabled] = useState<boolean>(true);
    const [promotedBtnsDisabled, setPromotedBtnsDisabled] = useState<boolean>(true);
    const [isRemoveDialogOpen, setIsRemoveDialogOpen] = useState<boolean>(false);
    const [isNotEditingQuestion, setIsNotEditingQuestion] = useState<boolean>(true);
    const [dialogContent, setDialogContent] = useState<string>();
    const [editedQuestion, setEditedQuestion] = useState<string>();
    const [isNotAllowedToSubmitQuestion, setIsNotAllowedToSubmitQuestion] = useState<boolean>(true);
    const [promotedItemsSearchFilter, setPromotedItemsSearchFilter] = useState<string>();
    const [notPromotedItemsSearchFilter, setNotPromotedItemsSearchFilter] = useState<string>();





    interface listItem {
        key: string;
        content: string;
        header?: string;
        promoted?: boolean;
    }

    useEffect(() => {
        // loadQuestions();
        updateQuestions();
    }, []);

    const listItems: listItem[] = allQuestions as listItem[];
    let promotedListItems: listItem[] = promotedQuestions as listItem[];
    let notPromotedListItems: listItem[] = notPromotedQuestions as listItem[];


    const updateQuestions = async () => {

        let listItems: listItem[] = [];
            loadQuestions().then((result: Question[]) => {
                for (let index = 0; index < result.length; index++) {
                    
                    const listItem: listItem = {
                        content: result[index].question,
                        key: result[index].RowKey,
                        header: result[index].author,
                        promoted: result[index].promoted
                    }

                    listItems.push(listItem);
                }

                setMyQuestions(listItems);

                promotedListItems = listItems.filter(item => item.promoted === true);
                setPromotedQuestions(promotedListItems);
                // console.log(promotedListItems);

                notPromotedListItems = listItems.filter(item => item.promoted === false);
                setNotPromotedQuestions(notPromotedListItems);
                // console.log(notPromotedListItems);
            });
    }

    const remoteQuestionFromArray = async () => {
        setNotPromotedQuestions(notPromotedQuestions?.filter((item) => item.key !== notPromotedQuestions![selectedNotPromotedListIndex!].key));
    }

    const loadQuestions = async () => {


        const myMeetingId: string = context?.meetingId as string;
        const myName: string = name as string;
        const fetchUrl: string = `/api/question?meetingId=${myMeetingId}&author=all`;
        const questionsList = await fetch(fetchUrl); 
        return questionsList.json();
    };

    const NotPromotedList = () => <List 
                                selectable
                                selectedIndex={selectedNotPromotedListIndex}
                                onSelectedIndexChange={(e, newProps: ListProps) => {
                                setSelectedNotPromotedListIndex(newProps.selectedIndex as number);
                                setNotPromotedBtnsDisabled(false);
                                // console.log(selectedListIndex);
                                }}
                                styles={{
                                    paddingTop: "0.625rem"
                                }}
                                items={notPromotedListItems.filter(listItem => {
                                    if (!notPromotedItemsSearchFilter) return true
                                    if (listItem.content.includes(notPromotedItemsSearchFilter) || listItem.header!.includes(notPromotedItemsSearchFilter)) {
                                        return true
                                    }
                                })} variables={{
                                    rootPadding: "0rem"
    }}/>

    const PromotedList = () => <List 
                                selectable
                                selectedIndex={selectedPromotedListIndex}
                                onSelectedIndexChange={(e, newProps: ListProps) => {
                                setSelectedPromotedListIndex(newProps.selectedIndex as number);
                                setPromotedBtnsDisabled(false);
                                // console.log(selectedListIndex);
                                }}
                                styles={{
                                    paddingTop: "0.625rem"
                                }}
                                items={promotedListItems.filter(listItem => {
                                    if (!promotedItemsSearchFilter) return true
                                    if (listItem.content.includes(promotedItemsSearchFilter) || listItem.header!.includes(promotedItemsSearchFilter)) {
                                        return true
                                    }
                                })} variables={{
                                    rootPadding: "0rem"
    }}/>

    // const updateQuestion = async () => {
    //     // console.log(question);
    //     setbtnUpdateLoading(true);
    //     // await updateTableQuestion();
    //     setEditedQuestion("");
    //     await updateQuestions();
    //     setbtnUpdateLoading(false);
    // };


    const handlePromoteBtnClicked = () => {
        updateTableQuestion("true", false);
    };

    const handleDemoteBtnClicked = () => {
        updateTableQuestion("false", true);
    };

    const handleRemoveBtnClicked = async () => {
        console.log("remove button clicked " + dialogContent);
        setIsRemoveDialogOpen(false);

        const res = await fetch(`/api/question?rowkey=${notPromotedListItems![selectedNotPromotedListIndex!].key}`, {
            method: "DELETE"
        });

        console.log(res.status);
        if (res.status === 200) {
            remoteQuestionFromArray();
        } else (console.log("an error happened"));

    };

    const handleNotPromotedCancelBtnClicked = async () => {
        console.log("cancel button clicked");
        setSelectedNotPromotedListIndex(-1);
        setNotPromotedBtnsDisabled(true);
    };

    const handlePromotedCancelBtnClicked = async () => {
        console.log("cancel button clicked");
        setSelectedPromotedListIndex(-1);
        setPromotedBtnsDisabled(true);
    };

    const updateTableQuestion = async (promoted: string, isPromotedList: boolean) => {

        let rowKey: string;

        if (isPromotedList) {
            rowKey = promotedListItems![selectedPromotedListIndex!].key as string;
        } else {
            rowKey = notPromotedListItems![selectedNotPromotedListIndex!].key as string;
        }

        // const rowkey: string = allQuestions![selectedPromotedListIndex!].key as string;
        const fetchUrl: string = `/api/question?rowkey=${rowKey!}&promoted=${promoted}`;

        const questionUpdateReponse = await fetch(fetchUrl, {
            method: "PATCH"
        });

        if (questionUpdateReponse.status === 200) {
            // To do add logic
            updateQuestions();
            !notPromotedBtnsDisabled && setNotPromotedBtnsDisabled(true);
            !promotedBtnsDisabled && setPromotedBtnsDisabled(true);
            setSelectedPromotedListIndex(-1);
            setSelectedNotPromotedListIndex(-1);

        } else (console.log("an error happened"));

        // console.log(questionUpdateReponse);

    };

    const handleNotPromotedQuestionsSearch = (event) => {
        setNotPromotedItemsSearchFilter(event.target.value);
    }

    const handlePromotedQuestionsSearch = (event) => {
        setPromotedItemsSearchFilter(event.target.value);
    }

    return (
        <>
        <Flex column padding="padding.medium">

        <Header
            as="h2"
            content="Meeting QnA - Organizer/Presenter"
            description={{
            content: "Use the options below to manage meeting questions",
            as: "span",
            }}
            styles={{
                paddingLeft: "1.250rem"
                // paddingBottom: "0.625rem"
        }}/>
        <RetryIcon onClick={updateQuestions} styles={{
                    position: "absolute",
                    right: "0",
                    marginTop: "1.625rem",
                    marginRight: "1.625rem",
                    top: "0",
                    cursor: "pointer"
        }}/>
        
    </Flex>

    <Flex gap="gap.small" padding="padding.medium">
        <Flex.Item size="size.half">
            <Segment>
                <Header as="h2" content="Not promoted questions" />
                <Flex gap="gap.small">
                    <Button disabled={notPromotedBtnsDisabled} onClick={handlePromoteBtnClicked} icon={<EyeIcon />} content="Promote" iconPosition="before" />
                    <Dialog
                        open={isRemoveDialogOpen}
                        onOpen={() => { setIsRemoveDialogOpen(true); setDialogContent(notPromotedListItems![selectedNotPromotedListIndex!].content) }}
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
                    trigger={<Button disabled={notPromotedBtnsDisabled} icon={<TrashCanIcon />} content="Remove" iconPosition="before" />}
                    />
                    <Button disabled={notPromotedBtnsDisabled} onClick={handleNotPromotedCancelBtnClicked} icon={<CloseIcon />} content="Cancel" iconPosition="before" />
                </Flex>
                <Flex gap="gap.small">
                    <Input fluid onChange={handleNotPromotedQuestionsSearch} icon={<SearchIcon />} placeholder="Search..." styles={{
                paddingTop: "0.625rem"
                }}/>
                </Flex>
                {notPromotedListItems ? <NotPromotedList /> : <Loader label="Loading not promoted questions" />}
            </Segment>
        </Flex.Item>
        <Flex.Item size="size.half">
            <Segment>
                <Header as="h2" content="Promoted questions" />
                <Flex gap="gap.small">
                <Button disabled={promotedBtnsDisabled} onClick={handleDemoteBtnClicked} icon={<EyeSlashIcon />} content="Demote" iconPosition="before" />
                <Button disabled={promotedBtnsDisabled} onClick={handlePromotedCancelBtnClicked} icon={<CloseIcon />} content="Cancel" iconPosition="before" />
                </Flex>
                <Flex gap="gap.small">
                    <Input onChange={handlePromotedQuestionsSearch} fluid icon={<SearchIcon />} placeholder="Search..." styles={{
                paddingTop: "0.625rem"
                }}/>
                </Flex>
                {promotedListItems ? <PromotedList /> : <Loader label="Loading promoted questions" />}
            </Segment>
        </Flex.Item>
    </Flex>
    </>
  )

};