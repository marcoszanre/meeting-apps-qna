import { Context } from "@microsoft/teams-js";
import * as React from "react";
import { FC } from "react";
import { Organizer } from "./Organizer";
import { Atendee } from "./Atendee";
import { OrganizerDetails } from "./OrganizerDetails";
import TaskContent from "./TaskContent";
import AtendeeDetails from "./AtendeeDetails";
import MeetingClosed from "./MeetingClosed";


interface IUIRouterProps {
  role: string;
  context: Context;
  name: string;
  active: boolean;
}

export const UIRouter: FC<IUIRouterProps> = ({ role, context, name, active }) => {

  const frameContext: string = context.frameContext as string;

  // handle default meeting State
  if (active === false && role !== "Organizer") {
    return <MeetingClosed />;
  }

  // handle default Task State
  if (frameContext === "task") {
    return <TaskContent context={context}/>;
  }

  // frameContext logic - "content"
  if (role === "Organizer" || role === "Presenter") {

    // content logic
    if (frameContext === "content") {
      // app is being loaded before/after meeting
      // console.log("app being loaded as Content");
      return (
        <Organizer context={context} name={name} />
      );
    } else {
      // app is being loaded inside the meeting as details
      // console.log("app being loaded as Details");
      return (
        <OrganizerDetails context={context} name={name}/>
      );
    }

  } else {

    // Attendee Content

    // content logic
    if (frameContext === "content") {
      // app is being loaded before/after meeting
        return (
          <Atendee context={context} name={name} />
        );
    } else {
      // app is being loaded inside the meeting as details
        return(
          <AtendeeDetails context={context} name={name} />
        );
    }
  }

};
