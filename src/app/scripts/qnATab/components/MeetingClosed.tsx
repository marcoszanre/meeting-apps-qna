import { Flex, Segment } from "@fluentui/react-northstar";
import * as React from "react";
import { FC } from "react";


const MeetingClosed: FC = () => {

    return (
        <>
        <Flex column hAlign="center" vAlign="center">
            <Flex.Item align="stretch">
            <Segment color="brand" content="This meeting is closed" inverted />
            </Flex.Item>
        </Flex>
        </>
    )

}

export default MeetingClosed;