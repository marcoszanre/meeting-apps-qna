import { Avatar, Button, Card, CardBody, CardHeader, DownloadIcon, EyeSlashIcon, Flex, Header, Loader, MoreIcon, RetryIcon, StarIcon } from "@fluentui/react-northstar";
import { Context } from "@microsoft/teams-js";
import * as React from "react";
import { FC, useEffect, useState } from "react";
import { Question } from "../../../services/tableService";
import CardsList from "./CardsList";
import TextExampleShorthand from "./Text";

interface IOrganizerDetailsProps {
    context: Context;
    name: string;
}

export const OrganizerDetails: FC<IOrganizerDetailsProps> = ({ context, name }) => {

    const [promotedQuestions, setPromotedQuestions] = useState<IListItem[]>();

    interface IListItem {
        key: string;
        content: string;
        header?: string;
        promoted?: boolean;
        Timestamp?: string;
        likedBy: number;
        asked?: boolean;
        askedWhen?: string;
    }

    let promotedListItems: IListItem[] = promotedQuestions as IListItem[];

    useEffect(() => {
        updateQuestions();
    }, []);

    const updateQuestions = async () => {

        const listItems: IListItem[] = [];
        loadQuestions().then((result: Question[]) => {
            for (let index = 0; index < result.length; index++) {

                const listItem: IListItem = {
                    content: result[index].question,
                    key: result[index].RowKey,
                    header: result[index].author,
                    promoted: result[index].promoted,
                    Timestamp: result[index].Timestamp,
                    likedBy: result[index].likedBy!,
                    asked: result[index].asked!,
                    askedWhen: result[index].askedWhen!
                };

                listItems.push(listItem);
            }

            promotedListItems = listItems.filter(item => item.promoted === true);
            setPromotedQuestions(promotedListItems);
            // console.log(promotedListItems);

        });
    };

    const loadQuestions = async () => {
        const myMeetingId: string = context?.meetingId as string;
        const fetchUrl: string = `/api/question?meetingId=${myMeetingId}&author=all`;
        const questionsList = await fetch(fetchUrl);
        return questionsList.json();
    };

    const sendBubble = async (listitem: IListItem) => {

        const activeQuestionData = {
            meetingid: context?.meetingId,
            question: listitem.content
        };

        const activeQuestionBody = JSON.stringify(activeQuestionData);

        await fetch("/api/activequestion", {
            method: "PATCH",
            headers: {
                "Content-Type": "application/json",
            },
            body: activeQuestionBody
        });

        const bubbleData = {
            chatId: context?.chatId,
            author: listitem.header,
            question: listitem.content
        };

        const body = JSON.stringify(bubbleData);

        const res = await fetch("/api/bubble", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: body
        });

        // if (res.status === 200) {
        //     remoteQuestionFromArray();
        // } else (console.log("an error happened"));

        // await fetch(`/api/bubble?chatId=${context.chatId}`);
        // alert(listitem.key);



        // UPDATED ASKED QUESTIONS

        const rowkey = listitem.key as string;

        const fetchUrl: string = `/api/askedquestion?rowkey=${rowkey}`;
        // const likeResponse = await (await fetch(fetchUrl)).json();
        // // console.log("meeting state is " + meetingStateResponse.meetingState);
        // // console.log(meetingStateResponse);

        // const likeData = {
        //     questionId: listitem.key,
        //     userID: context.userObjectId,
        // };

        // const body = JSON.stringify(likeData);

        // const res = await fetch("/api/like", {
        //     method: "POST",
        //     headers: {
        //         "Content-Type": "application/json",
        //     },
        //     body: body
        // });

        // // console.log(res.status);

        // // check if user has already liked or not
        // !likeResponse.like ? listitem.likedBy! += 1 : listitem.likedBy! -= 1;

        // setPromotedQuestions(
        //     promotedQuestions!.map(item =>
        //         item.key === listitem.key
        //         ? {...item, likedBy : listitem.likedBy!}
        //         : item
        // ));


    };

    return (
        <>
        <Flex column padding="padding.medium">
        <Header
            as="h3"
            content="Meeting questions"
            description={{
            content: "Organizer/Presenter",
            as: "span",
            }}
            styles={{
                paddingLeft: "0.625rem",
                paddingBottom: "0.625rem"
        }}/>

        <RetryIcon title="Refresh Questions" onClick={updateQuestions} styles={{
                    position: "absolute",
                    right: "0",
                    marginTop: "2.250rem",
                    marginRight: "1.250rem",
                    top: "0",
                    cursor: "pointer"
        }}/>

        {promotedQuestions ? promotedQuestions.map((listitem: IListItem) =>

            <Flex column gap="gap.medium">
            <Card fluid key={listitem.key}>
                <Card.Header>
                    <Flex gap="gap.small">
                        <Avatar name={listitem.header!}/>
                        <Flex column>
                            <TextExampleShorthand content={listitem.header!}/>
                            <TextExampleShorthand content={listitem.Timestamp!.split("T")[0]}/>
                        </Flex>
                    </Flex>
                </Card.Header>
                <Card.Body>
                    <Flex column gap="gap.small">
                        <TextExampleShorthand content={listitem.content}/>
                    </Flex>
                </Card.Body>
                <Card.Footer>
                    <Flex space="between">
                        <Button primary onClick={() => sendBubble(listitem)} content="Promote" />
                        <Flex vAlign="center">
                            <TextExampleShorthand content={`${listitem.likedBy} likes`} />
                            <EyeSlashIcon outline={!listitem.asked} title={listitem.asked ? "Question Asked" : "Question Not Asked"} styles={{
                                        cursor: "pointer"
                            }}/>
                        </Flex>
                    </Flex>
                </Card.Footer>
            </Card>
            </Flex>
            )
            : <Loader label="Loading promoted questions" />}

        </Flex>

        </>
  );

};
