import { Avatar, Button, Card, CardBody, CardHeader, DownloadIcon, Flex, Header, Loader, MoreIcon, RetryIcon, StarIcon } from "@fluentui/react-northstar";
import { Context } from "@microsoft/teams-js";
import * as React from "react";
import { FC, useEffect, useState } from "react";
import { Question } from "../../../services/tableService";
import CardsList from "./CardsList";
import TextExampleShorthand from "./Text";

type OrganizerDetailsProps = {
    context: Context,
    name: string,
}

export const OrganizerDetails: FC<OrganizerDetailsProps> = ({ context, name }) => {

    const [promotedQuestions, setPromotedQuestions] = useState<listItem[]>();

    interface listItem {
        key: string;
        content: string;
        header?: string;
        promoted?: boolean;
        Timestamp?: string;
        likedBy: number;
    }

    let promotedListItems: listItem[] = promotedQuestions as listItem[];

    useEffect(() => {
        updateQuestions();
    }, []);

    const updateQuestions = async () => {

        let listItems: listItem[] = [];
            loadQuestions().then((result: Question[]) => {
                for (let index = 0; index < result.length; index++) {
                    
                    const listItem: listItem = {
                        content: result[index].question,
                        key: result[index].RowKey,
                        header: result[index].author,
                        promoted: result[index].promoted,
                        Timestamp: result[index].Timestamp,
                        likedBy: result[index].likedBy!
                    }

                    listItems.push(listItem);
                }

                promotedListItems = listItems.filter(item => item.promoted === true);
                setPromotedQuestions(promotedListItems);
                // console.log(promotedListItems);

            });
    }

    // const cards = promotedQuestions.map((listitem: listItem) => 
    //     <Flex column gap="gap.medium">
    //     <Card fluid key={listitem.key}>
    //         <Card.Header>
    //             <Flex gap="gap.small">
    //                 <Avatar name="20"/>
    //                 <Flex column>
    //                     <TextExampleShorthand content={listitem.header!}/>
    //                     <TextExampleShorthand content={listitem.Timestamp!.split("T")[0]}/>
    //                 </Flex>
    //             </Flex>
    //         </Card.Header>
    //         <Card.Body>
    //             <Flex column gap="gap.small">
    //                 <TextExampleShorthand content={listitem.content}/>
    //             </Flex>
    //         </Card.Body>
    //         <Card.Footer>
    //             <Flex space="between">
    //                 <Button content="Promote" />
    //             </Flex>
    //         </Card.Footer>
    //     </Card>
    //     </Flex>
    // );
    


    const loadQuestions = async () => {
        const myMeetingId: string = context?.meetingId as string;
        const fetchUrl: string = `/api/question?meetingId=${myMeetingId}&author=all`;
        const questionsList = await fetch(fetchUrl); 
        return questionsList.json();
    };

    const sendBubble = async (listitem: listItem) => {

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
    }

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

        <RetryIcon onClick={updateQuestions} styles={{
                    position: "absolute",
                    right: "0",
                    marginTop: "2.250rem",
                    marginRight: "1.250rem",
                    top: "0",
                    cursor: "pointer"
        }}/>

        {promotedQuestions ? promotedQuestions.map((listitem: listItem) => 
            
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
                        </Flex>
                    </Flex>
                </Card.Footer>
            </Card>
            </Flex>
            )
            : <Loader label="Loading promoted questions" />}
        
        </Flex>

        </>
  )

};