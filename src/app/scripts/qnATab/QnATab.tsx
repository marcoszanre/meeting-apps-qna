import * as React from "react";
import { Provider, Flex, Button, Header, Segment, Alert, TextArea, List, ButtonGroup, Loader, ListProps, Dialog, Input, Table } from "@fluentui/react-northstar";
import { CloseIcon, EditIcon } from "@fluentui/react-icons-northstar";
import { useState, useEffect } from "react";
import { useTeams } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import jwt_decode from "jwt-decode";
import { Question, setMeetingState } from "../../services/tableService";
import { UIRouter } from "./components/UIRouter";

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
    const [isMeetingStateActive, setIsMeetingStateActive] = useState<boolean>();


    useEffect(() => {
        if (inTeams === true) {

            microsoftTeams.authentication.getAuthToken({
                successCallback: (token: string) => {
                    // console.log(token);
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
            // console.log(context);
        }
    }, [context]);

    // call load questions when once name is set
    useEffect(() => {
        if (name && context) {
            // updateQuestions();
            getParticipantRole();
            createMeetingState();
        }
    }, [name]);

    const createMeetingState = async () => {

        const fetchUrl: string = `/api/meetingstate?meetingid=${context?.meetingId!}`;
        const meetingStateResponse = await (await fetch(fetchUrl)).json();
        // console.log("meeting state is " + meetingStateResponse.meetingState);
        // console.log(meetingStateResponse);

        if (meetingStateResponse.meetingState === "not found") {

            setIsMeetingStateActive(true);

            const meetingData = {
                meetingid: context?.meetingId as string,
                active: true as boolean,
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

            // console.log(res.status);

            // console.log("no meeting state found");

        } else {

            setIsMeetingStateActive(meetingStateResponse.meetingState);
            // console.log("meeting state is " + meetingStateResponse.meetingState);
        }

    };

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
    };

    return (
        <Provider theme={theme}>

            {/* {userRole ? <UIRouter role={userRole} context={context!} name={name!}/> : <Loader label="Loading details" />} */}
            {userRole && <UIRouter role={userRole} context={context!} name={name!} active={isMeetingStateActive!} />}

        </Provider>
    );
};
