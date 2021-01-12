import * as React from "react";
import { Provider, Flex, Text, Header } from "@fluentui/react-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * Implementation of QnA remove page
 */
export const QnATabRemove = () => {

    const [{ inTeams, theme }] = useTeams();

    useEffect(() => {
        if (inTeams === true) {
            microsoftTeams.appInitialization.notifySuccess();
        }
    }, [inTeams]);

    return (
        <Provider theme={theme}>
            <Flex fill={true}>
                <Flex.Item>
                    <div>
                        <Header content="You're about to remove your tab..." />
                        <Text content="Thanks for using QnA!" />
                    </div>
                </Flex.Item>
            </Flex>
        </Provider>
    );
};
