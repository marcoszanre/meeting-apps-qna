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



    interface listItem {
        key: string;
        content: string;
        header?: string;
        promoted?: boolean;
        Timestamp?: string;
        likedBy?: string;
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
                        likedBy: result[index].likedBy
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
        const reactionCount = listitem.likedBy!.split(",").length;
        hasNotReacted ? setReactionCount(reactionCount + 1) : setReactionCount(reactionCount - 1);
        setHasNotReacted(!hasNotReacted);
        alert(listitem.likedBy!);
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
                        <TextExampleShorthand content={listitem.likedBy!}/>
                        <TextExampleShorthand content={listitem.likedBy!.split(",")[0] } />
                    </Flex>
                </Card.Body>
                <Card.Footer>
                    <Flex space="between">
                        <Reaction onClick={() => handleReactionClick(listitem)} icon={<LikeIcon outline={hasNotReacted} />} content={listitem.likedBy!.split(",").length } />
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