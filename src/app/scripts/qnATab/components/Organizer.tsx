import { Flex, Header, Segment, Input, Dialog, CloseIcon, Button, Loader, TextArea, List, ListProps, TrashCanIcon, EditIcon, AcceptIcon, EyeIcon, EyeSlashIcon, SearchIcon, RetryIcon, BanIcon, CallRecordingIcon } from "@fluentui/react-northstar";
import { Context } from "@microsoft/teams-js";
import { PowerBIEmbed } from "powerbi-client-react";
import { models, Report, Embed, IEmbedConfiguration, service, Page } from 'powerbi-client';
import * as React from "react";
import { useState, useEffect } from "react";
import { Question } from "../../../services/tableService";

interface IOrganizerProps {
    context: Context;
    name: string;
    teamsAccessToken: string;
    isDefaultMeetingActive: boolean;
}

export const Organizer: React.FC<IOrganizerProps> = ({ context, name, teamsAccessToken, isDefaultMeetingActive }) => {

    const [question, setQuestion] = useState<string>();
    const [allQuestions, setMyQuestions] = useState<IListItem[]>();
    const [promotedQuestions, setPromotedQuestions] = useState<IListItem[]>();
    const [notPromotedQuestions, setNotPromotedQuestions] = useState<IListItem[]>();
    const [btnLoading, setBtnLoading] = useState<boolean>(false);
    const [btnUpdateLoading, setbtnUpdateLoading] = useState<boolean>(false);
    const [selectedPromotedListIndex, setSelectedPromotedListIndex] = useState<number>(-1);
    const [selectedNotPromotedListIndex, setSelectedNotPromotedListIndex] = useState<number>(-1);
    const [removeBtnDisabled, setRemoveBtnDisabled] = useState<boolean>(true);
    const [notPromotedBtnsDisabled, setNotPromotedBtnsDisabled] = useState<boolean>(true);
    const [promotedBtnsDisabled, setPromotedBtnsDisabled] = useState<boolean>(true);
    const [isRemoveDialogOpen, setIsRemoveDialogOpen] = useState<boolean>(false);
    const [isCloseMeetingDialogOpen, setIsCloseMeetingDialogOpen] = useState<boolean>(false);
    const [isNotEditingQuestion, setIsNotEditingQuestion] = useState<boolean>(true);
    const [isMeetingStateActive, setIsMeetingStateActive] = useState<boolean>(isDefaultMeetingActive);
    const [dialogContent, setDialogContent] = useState<string>();
    const [editedQuestion, setEditedQuestion] = useState<string>();
    const [isNotAllowedToSubmitQuestion, setIsNotAllowedToSubmitQuestion] = useState<boolean>(true);
    const [promotedItemsSearchFilter, setPromotedItemsSearchFilter] = useState<string>();
    const [notPromotedItemsSearchFilter, setNotPromotedItemsSearchFilter] = useState<string>();
    const [accessToken, setAccessToken] = useState<string>();


    interface IListItem {
        key: string;
        content: string;
        header?: string;
        promoted?: boolean;
    }

    useEffect(() => {
        // loadQuestions();
        // updateQuestions();
        initializeQuestions();
        // loadMeetingState();
        (!isDefaultMeetingActive && allQuestions?.length! > 0) && initializePowerBI();
    }, []);

    const initializePowerBI = async () => {
        updatePowerBIReactClass();
        await loadPowerBIAccessToken();
    }

    const initializeQuestions = async () => {
        await updateQuestions();
    }

    const listItems: IListItem[] = allQuestions as IListItem[];
    let promotedListItems: IListItem[] = promotedQuestions as IListItem[];
    let notPromotedListItems: IListItem[] = notPromotedQuestions as IListItem[];

    const updatePowerBIReactClass = () => {
        let elm: HTMLElement;
        elm = document.querySelector<HTMLElement>(".powerBIClass")!;
        // console.log(elm);
        elm!.style.height = "24rem";
    }

    const updateQuestions = async () => {

        const listItems: IListItem[] = [];
        loadQuestions().then((result: Question[]) => {
            for (let index = 0; index < result.length; index++) {

                const listItem: IListItem = {
                    content: result[index].question,
                    key: result[index].RowKey,
                    header: result[index].author,
                    promoted: result[index].promoted
                };

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
    };

    const remoteQuestionFromArray = async () => {
        setNotPromotedQuestions(notPromotedQuestions?.filter((item) => item.key !== notPromotedQuestions![selectedNotPromotedListIndex!].key));
    };

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
                                    if (!notPromotedItemsSearchFilter) { return true; }
                                    if (listItem.content.includes(notPromotedItemsSearchFilter) || listItem.header!.includes(notPromotedItemsSearchFilter)) {
                                        return true;
                                    }
                                })} variables={{
                                    rootPadding: "0rem"
    }}/>;

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
                                    if (!promotedItemsSearchFilter) { return true; }
                                    if (listItem.content.includes(promotedItemsSearchFilter) || listItem.header!.includes(promotedItemsSearchFilter)) {
                                        return true;
                                    }
                                })} variables={{
                                    rootPadding: "0rem"
    }}/>;

    const handlePromoteBtnClicked = () => {
        updateTableQuestion("true", false);
    };

    const handleDemoteBtnClicked = () => {
        updateTableQuestion("false", true);
    };

    const handleCloseMeetingBtnClicked = async (state: boolean) => {
        const meetingData = {
            meetingid: context?.meetingId,
            active: state,
        };

        const body = JSON.stringify(meetingData);
        // console.log(body);

        const res = await fetch("/api/meetingstate", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: body
        });

        // console.log("the meeting has been closed");
        setIsCloseMeetingDialogOpen(false);

        isMeetingStateActive && await loadPowerBIAccessToken();
        setIsMeetingStateActive(state);
        // !state && updatePowerBIReactClass();
        (!state && allQuestions?.length! > 0) && updatePowerBIReactClass();

    };

    const loadPowerBIAccessToken = async () => {

        const fetchUrl: string = `/api/powerbiaccesstoken?token=${teamsAccessToken}`;
        const accessTokenReponse = await (await fetch(fetchUrl)).json();
        // console.log(questionUpdateReponse.meetingState);
        setAccessToken(accessTokenReponse.accessToken);
        
    }


    const loadMeetingState = async () => {

        const fetchUrl: string = `/api/meetingstate?meetingid=${context.meetingId}`;

        const questionUpdateReponse = await (await fetch(fetchUrl)).json();
        // console.log(questionUpdateReponse.meetingState);

        setIsMeetingStateActive(questionUpdateReponse.meetingState);

    };

    const handleRemoveBtnClicked = async () => {
        // console.log("remove button clicked " + dialogContent);
        setIsRemoveDialogOpen(false);

        const res = await fetch(`/api/question?rowkey=${notPromotedListItems![selectedNotPromotedListIndex!].key}`, {
            method: "DELETE"
        });

        // console.log(res.status);
        if (res.status === 200) {
            remoteQuestionFromArray();
        } else {
            // console.log("an error happened");
        }

    };

    const handleNotPromotedCancelBtnClicked = async () => {
        // console.log("cancel button clicked");
        setSelectedNotPromotedListIndex(-1);
        setNotPromotedBtnsDisabled(true);
    };

    const handlePromotedCancelBtnClicked = async () => {
        // console.log("cancel button clicked");
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

        } else {
            // console.log("an error happened");
        }

        // console.log(questionUpdateReponse);

    };

    const handleNotPromotedQuestionsSearch = (event) => {
        setNotPromotedItemsSearchFilter(event.target.value);
    };

    const handlePromotedQuestionsSearch = (event) => {
        setPromotedItemsSearchFilter(event.target.value);
    };

    const basicFilter: models.IBasicFilter = {
        $schema: "http://powerbi.com/product/schema#basic",
        target: {
          table: "questionsTable",
          column: "meetingid"
        },
        operator: "In",
        values: [context.meetingId!],
        filterType: 1, // pbi.models.FilterType.BasicFilter
        requireSingleSelection: true // Limits selection of values to one.
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
        <RetryIcon aria-label="Refresh Questions"
            onClick={updateQuestions} styles={{
                    position: "absolute",
                    right: "0",
                    marginTop: "1.625rem",
                    marginRight: "1.625rem",
                    top: "0",
                    cursor: "pointer"
        }}/>
        <Dialog
                        open={isCloseMeetingDialogOpen}
                        onOpen={() => setIsCloseMeetingDialogOpen(true)}
                        onCancel={() => setIsCloseMeetingDialogOpen(false)}
                        onConfirm={() => handleCloseMeetingBtnClicked(!isMeetingStateActive)}
                        confirmButton="Confirm"
                        cancelButton="Cancel"
                        content={context.meetingId}
                        header={isMeetingStateActive ? "Are you sure you want to close this meeting?" : "Are you sure you want to reopen this meeting?"}
                        headerAction={{
                            icon: <CloseIcon />,
                            title: "Close",
                            onClick: () => setIsCloseMeetingDialogOpen(false),
                    }}
                    trigger={isMeetingStateActive ?
                        <BanIcon aria-label="Close Meeting" styles={{
                            position: "absolute",
                            right: "0",
                            marginTop: "1.625rem",
                            marginRight: "3.625rem",
                            top: "0",
                            cursor: "pointer"
                    }} /> : <CallRecordingIcon aria-label="Open Meeting" styles={{
                        position: "absolute",
                        right: "0",
                        marginTop: "1.625rem",
                        marginRight: "3.625rem",
                        top: "0",
                        cursor: "pointer"
                }} />
                }
        />

    </Flex>

    { (!isMeetingStateActive && allQuestions?.length! > 0) && <Flex styles={{ height: "28rem" }} column padding="padding.medium">
            <>
            <Segment color="brand" content="This report updates daily every hour from 9AM to 6PM." />
            <PowerBIEmbed cssClassName="powerBIClass"
                embedConfig = {{
                    type: "report",   // Supported types: report, dashboard, tile, visual and qna
                    id: "5e67a94d-02e4-45d5-a6a2-25e5c39b43f3",
                    settings: {
                        filterPaneEnabled: false,
                        navContentPaneEnabled: false,
                        persistentFiltersEnabled: false,
                    },
                    embedUrl: "https://app.powerbi.com/reportEmbed?reportId=5e67a94d-02e4-45d5-a6a2-25e5c39b43f3&groupId=01855087-cf91-4f29-90f0-04b595b49cdf&config=eyJjbHVzdGVyVXJsIjoiaHR0cHM6Ly9XQUJJLVVTLUNFTlRSQUwtQS1QUklNQVJZLXJlZGlyZWN0LmFuYWx5c2lzLndpbmRvd3MubmV0IiwiZW1iZWRGZWF0dXJlcyI6eyJtb2Rlcm5FbWJlZCI6dHJ1ZX19",
                    accessToken: accessToken,    // Keep as empty string, null or undefined
                    // accessToken: "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6IjVPZjlQNUY5Z0NDd0NtRjJCT0hIeEREUS1EayIsImtpZCI6IjVPZjlQNUY5Z0NDd0NtRjJCT0hIeEREUS1EayJ9.eyJhdWQiOiJodHRwczovL2FuYWx5c2lzLndpbmRvd3MubmV0L3Bvd2VyYmkvYXBpIiwiaXNzIjoiaHR0cHM6Ly9zdHMud2luZG93cy5uZXQvMjk2OTU2YWEtOWNmNC00MmQ5LTgwZjUtNGM0ZTRmOGM0ZmYxLyIsImlhdCI6MTYxMDAzOTQ1MiwibmJmIjoxNjEwMDM5NDUyLCJleHAiOjE2MTAwNDMzNTIsImFjY3QiOjAsImFjciI6IjEiLCJhaW8iOiJFMkpnWVBqOWU3dE9rTm9ldzVBK2xqZHQ2a3MzZEJaT3l1R2Zlb1M3aEUva2xEY2Yxd01BIiwiYW1yIjpbInB3ZCJdLCJhcHBpZCI6ImM4NWMxNzM1LTMyZDYtNDAzOC1iY2IzLWI3NWI0NzZmNDlmZSIsImFwcGlkYWNyIjoiMSIsImZhbWlseV9uYW1lIjoiQWRtaW5pc3RyYXRvciIsImdpdmVuX25hbWUiOiJNT0QiLCJpcGFkZHIiOiIxODcuMTAxLjEwNC4xNCIsIm5hbWUiOiJNT0QgQWRtaW5pc3RyYXRvciIsIm9pZCI6IjkzZjZkMGNkLTViZjctNDhkNy1hYzI1LWQ3ODQ5ZjdhMzEwMiIsInB1aWQiOiIxMDAzMjAwMDk0NjQ1MUI1IiwicmgiOiIwLkFBQUFxbFpwS2ZTYzJVS0E5VXhPVDR4UDhUVVhYTWpXTWpoQXZMTzNXMGR2U2Y1WEFFUS4iLCJzY3AiOiJSZXBvcnQuUmVhZC5BbGwiLCJzdWIiOiJlM1RmWE1YR1lSQUwzM1dELTVpcmMteFU2Z3U0Y1JVMG9wcjVnVTVZOU5FIiwidGlkIjoiMjk2OTU2YWEtOWNmNC00MmQ5LTgwZjUtNGM0ZTRmOGM0ZmYxIiwidW5pcXVlX25hbWUiOiJhZG1pbkBNMzY1eDE2NTc1My5vbm1pY3Jvc29mdC5jb20iLCJ1cG4iOiJhZG1pbkBNMzY1eDE2NTc1My5vbm1pY3Jvc29mdC5jb20iLCJ1dGkiOiJTQW0xLTNRcU9VYTB6MkhwTjk4WkFRIiwidmVyIjoiMS4wIiwid2lkcyI6WyJlNmQxYTIzYS1kYTExLTRiZTQtOTU3MC1iZWZjODZkMDY3YTciLCI2MmU5MDM5NC02OWY1LTQyMzctOTE5MC0wMTIxNzcxNDVlMTAiLCJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXX0.Uc0vd_1r_KIV7gyOGhy12nbRlN5yIfdXx8qFDD9brVb3g-IryYv-qUwf8xfRQMHj_gVy9r4VDCT8itV0r2PM3Z5WSUv3t9mp5VRJ3gdbWny1CvK7HE8B7IpfAE5W6p0yGKazEEziP-oZL5g-Rifp2XUEdibC-iE5khXYgbMJO2mvz5_PXtPyjyT-l2s5bi0TADVqTfzO9sLjWCnaBS2Z2YuftvkEDIhQRQMSY19haTEknke6ooSR4lm4q_nPtNdAIHiv4jKBx3Csuk0ZEbO4Niy_ect1QRwrTnM3QxXoqxbQA2xc00LzwvK8uf9VCneM8mRdgbwobYRvQeSHzbVXlA",    // Keep as empty string, null or undefined
                    tokenType: models.TokenType.Aad,
                    pageName: "Home",
                    filters: [ basicFilter ]
                }}
            />
            </>
    </Flex>
    }

    <Flex gap="gap.small" padding="padding.medium">
        <Flex.Item size="size.half">
            <Segment>
                <Header as="h2" content="Not promoted questions" />
                <Flex gap="gap.small">
                    <Button disabled={notPromotedBtnsDisabled} onClick={handlePromoteBtnClicked} icon={<EyeIcon />} content="Promote" iconPosition="before" />
                    <Dialog
                        open={isRemoveDialogOpen}
                        onOpen={() => { setIsRemoveDialogOpen(true); setDialogContent(notPromotedListItems![selectedNotPromotedListIndex!].content); }}
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
  );

};
