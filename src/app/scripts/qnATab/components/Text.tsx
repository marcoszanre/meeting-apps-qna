import * as React from "react";
import { Text } from "@fluentui/react-northstar";
import { FC } from "react";

interface ITextExampleShorthandProps {
    content: string;
}

const TextExampleShorthand: FC<ITextExampleShorthandProps> = ({ content }) => (
  <Text content={content} />
);

export default TextExampleShorthand;
