import * as React from "react";
import { Avatar, Button, Card, Flex, Loader, Text } from "@fluentui/react-northstar";
import { FC, useEffect } from "react";
import TextExampleShorthand from "./Text";

interface ICardsListProps {
    questions: IListItem[];
}

interface IListItem {
    key: string;
    content: string;
    header?: string;
    promoted?: boolean;
    Timestamp?: string;
}

const CardsList: FC<ICardsListProps> = ({ questions }) => {


    const cards = questions.map((listitem: IListItem) =>
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

    return (
        <>
        {cards}
        </>
    );
};

export default CardsList;
