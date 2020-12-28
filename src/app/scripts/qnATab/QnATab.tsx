import * as React from "react";
import { Provider, Flex, Button, Header, Segment, Alert, TextArea, List, ButtonGroup, Loader, ListProps, Dialog, Input, Table } from "@fluentui/react-northstar";
import { CloseIcon, EditIcon } from "@fluentui/react-icons-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwt_decode from "jwt-decode";
import { Question } from "../../services/tableService";
import { Details } from "./components/Details";

/**
 * Implementation of the QnA content page
 */
export const QnATab = () => {

    const [{ inTeams, theme, context }] = useTeams();
    const [entityId, setEntityId] = useState<string | undefined>();
    const [name, setName] = useState<string>();
    const [error, setError] = useState<string>();
    // const [myQuestions, setMyQuestions] = useState<listItem[]>();
    const [userRole, setUserRole] = useState<string>();


    useEffect(() => {
        if (inTeams === true) {

            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    console.log(token);
                    const decoded: { [key: string]: any; } = jwt_decode(token) as { [key: string]: any; };
                    setName(decoded!.name);
                    // microsoftTeams.appInitialization.notifySuccess();
                },
                failureCallback: (message: string) => {
                    setError(message);
                    microsoftTeams.appInitialization.notifyFailure({
                        reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                        message
                    });
                },
                resources: [process.env.QNA_APP_URI as string]
            });
        } else {
            setEntityId("Not in Microsoft Teams");
        }
    }, [inTeams]);

    useEffect(() => {
        if (context) {
            setEntityId(context.entityId);
            console.log(context);
        }
    }, [context]);

    // call load questions when once name is set
    useEffect(() => {
        if (name && context) {
            // updateQuestions();
            getParticipantRole();
        }
    }, [name]);

    // useEffect(() => {
    //     if (userRole) {
    //         loadUI();
    //     }
    // }, [userRole]);

    const getParticipantRole = async () => {

        const res = await fetch(`/api/role?meetingId=${context?.meetingId}&userId=${context?.userObjectId}`);
        const json = await res.json();
        // console.log(json);
        setUserRole(json.role);
        microsoftTeams.appInitialization.notifySuccess();
    }

    // const loadUI = () => {
    //     if (userRole === "Organizer") {
    //         // load organizer UI
    //         console.log("You're an organizer");
    //     } else {
    //         // load user UI
    //         console.log("You're an attendee or presenter");
    //     }
    // }

    
    return (
        <Provider theme={theme}>

            {userRole ? <Details role={userRole} context={context!} name={name!}/> : <Loader label="Loading details" />}

        </Provider>
    );
};
