import { Context } from "@microsoft/teams-js";
import * as React from "react";
import { FC } from "react";
import { Organizer } from "./Organizer";
import { Atendee } from "./Atendee";


type DetailsProps = {
  role: string,
  context: Context,
  name: string,
}

export const Details: FC<DetailsProps> = ({ role, context, name }) => {

  if (role === "Organizer") {
    return (
      <Organizer context={context} name={name} />
    )
  } else {
    return (
      <Atendee context={context} name={name} />
    )
  }

};