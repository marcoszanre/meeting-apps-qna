import * as React from "react";
import { Text } from '@fluentui/react-northstar';
import { FC } from "react";

type TextExampleShorthandProps = {
    content: string;
}

const TextExampleShorthand: FC<TextExampleShorthandProps> = ({ content }) => (
  <Text content={content} />
)

export default TextExampleShorthand;