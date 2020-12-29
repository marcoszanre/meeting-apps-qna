import * as React from "react";
import { Avatar, Button, Card, Flex, Loader, Text } from '@fluentui/react-northstar';
import { FC, useEffect } from "react";
import TextExampleShorthand from "./Text";

type CardsListProps = {
    questions: listItem[];
}

interface listItem {
    key: string;
    content: string;
    header?: string;
    promoted?: boolean;
    Timestamp?: string;
}

const CardsList: FC<CardsListProps> = ({ questions }) => {

    // const logQuestions = () => {
    //     console.log(questions);
    // }

    // const cards = questions.map((listitem) => 
    //     <li>{listitem.content}</li>
    // );


    const cards = questions.map((listitem: listItem) => 
        <Flex column gap="gap.medium">
        <Card fluid key={listitem.key}>
            <Card.Header>
                <Flex gap="gap.small">
                    <Avatar name="20"/>
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
                    <Button content="Promote" />
                </Flex>
            </Card.Footer>
        </Card>
        </Flex>
    );

    return(
    <>
    {cards}
    {/* <ul>{cards}</ul> */}
            {/* <Flex column gap="gap.small">
                <Card fluid>
                    <Card.Header>
                        <Flex gap="gap.small">
                            <Avatar name="20"/>
                            <Flex column>
                                <TextExampleShorthand content="{listitem.header!}"/>
                                <TextExampleShorthand content="20/20/2020"/>
                            </Flex>
                        </Flex>
                    </Card.Header>
                    <Card.Body>
                        <Flex column gap="gap.small">
                            <TextExampleShorthand content="{listitem.content}"/>
                        </Flex>
                    </Card.Body>
                    <Card.Footer>
                        <Flex space="between">
                            <Button onClick={logQuestions} content="Promote" />
                        </Flex>
                    </Card.Footer>
                </Card>
                </Flex>
                <Flex column gap="gap.small">
                <Card fluid>
                    <Card.Header>
                        <Flex gap="gap.small">
                            <Avatar name="20"/>
                            <Flex column>
                                <TextExampleShorthand content="{listitem.header!}"/>
                                <TextExampleShorthand content="20/20/2020"/>
                            </Flex>
                        </Flex>
                    </Card.Header>
                    <Card.Body>
                        <Flex column gap="gap.small">
                            <TextExampleShorthand content="{listitem.content}"/>
                        </Flex>
                    </Card.Body>
                    <Card.Footer>
                        <Flex space="between">
                            <Button onClick={() => console.log("cliekd")} content="Promote" />
                        </Flex>
                    </Card.Footer>
                </Card>
                </Flex> */}
                </>
            )

//     return (
//         <Flex column gap="gap.small">
//             {cards}
//         </Flex>
//     )
// }
}

export default CardsList;