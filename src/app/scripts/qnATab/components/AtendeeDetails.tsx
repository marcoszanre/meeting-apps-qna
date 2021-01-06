import { FC, useState } from "react";
import * as React from "react";
import { Context } from "@microsoft/teams-js";
import { Flex, Header, RetryIcon, Card, Avatar, Button, StarIcon, Loader, LikeIcon, Reaction } from "@fluentui/react-northstar";
import { Question } from "../../../services/tableService";
import TextExampleShorthand from "./Text";

type AtendeeDetailsProps = {
    context: Context,
    name: string,
}

const AtendeeDetails: FC<AtendeeDetailsProps> = ({ context, name }) => {

    const [promotedQuestions, setPromotedQuestions] = useState<listItem[]>();
    const [hasNotReacted, setHasNotReacted] = useState<boolean>(true);
    const [reactionCount, setReactionCount] = useState<number>(10);
    const [log, setLog] = useState<string>("");



    interface listItem {
        key: string;
        content: string;
        header?: string;
        promoted?: boolean;
        Timestamp?: string;
        likedBy: number;
    }

    let promotedListItems: listItem[] = promotedQuestions as listItem[];

    React.useEffect(() => {
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


    const loadQuestions = async () => {
        const myMeetingId: string = context?.meetingId as string;
        const fetchUrl: string = `/api/question?meetingId=${myMeetingId}&author=all`;
        const questionsList = await fetch(fetchUrl); 
        return questionsList.json();
    };


    const handleReactionClick = async (listitem: listItem) => {

        const key = listitem.key as string;
        const userId = context.userObjectId as string;
        // alert(key);
        // alert(userId);
        
        const fetchUrl: string = `/api/like?questionId=${key}&userID=${userId}`;
        const likeResponse = await (await fetch(fetchUrl)).json();
        // console.log("meeting state is " + meetingStateResponse.meetingState);
        // console.log(meetingStateResponse);
       
        const likeData = {
            questionId: listitem.key,
            userID: context.userObjectId,
        };

        const body = JSON.stringify(likeData);

        const res = await fetch("/api/like", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
            },
            body: body
        });

        // console.log(res.status);

        // check if user has already liked or not
        !likeResponse.like ? listitem.likedBy! += 1 : listitem.likedBy! -= 1;

        setPromotedQuestions(
            promotedQuestions!.map(item => 
                item.key === listitem.key 
                ? {...item, likedBy : listitem.likedBy!} 
                : item 
        ));

        // setHasNotReacted(!hasNotReacted);
        // alert(listitem.likedBy!);
    };

    return (
        <>
        <Flex column padding="padding.medium">
        <Header
            as="h3"
            content="Meeting questions"
            description={{
            content: "Atendee",
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

        <TextExampleShorthand content={log}/>


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
                        <Reaction onClick={() => handleReactionClick(listitem)} icon={<LikeIcon outline={ false } />} content={ listitem.likedBy } />
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

export default AtendeeDetails;