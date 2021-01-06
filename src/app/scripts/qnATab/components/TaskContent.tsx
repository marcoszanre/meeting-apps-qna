import { Button, Flex, Header } from "@fluentui/react-northstar";
import { Context } from "@microsoft/teams-js";
import * as React from "react";
import { FC, useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";


interface ITaskContentProps {
    context: Context;
}

const TaskContent: FC<ITaskContentProps> = ({ context }) => {

    const [activeQuestion, setActiveQuestion] = useState<string>();

    const retrieveActiveQuestion = async () => {
        const meetingid = context.meetingId as string;
        const res = await fetch(`/api/activequestion?meetingid=${meetingid}`);
        const json = await res.json();
        setActiveQuestion(json.activeQuestion);
    };

    // call load active questions
    useEffect(() => {
        retrieveActiveQuestion();
    }, []);

    return (
        <>
        <Flex column space="between">
        <Header
            as="h2"
            content={activeQuestion}
            align="center"
            styles={{
                // paddingLeft: "0.650rem",
                // paddingBottom: "0.625rem"
        }}/>
        <Button onClick={() => microsoftTeams.tasks.submitTask()} content="OK" primary styles={{
                    position: "absolute",
                    right: "0",
                    marginBottom: "0.625rem",
                    marginRight: "0.625rem",
                    bottom: "0",
                    cursor: "pointer"
        }}/>
        </Flex>
        </>
    );

};

export default TaskContent;
