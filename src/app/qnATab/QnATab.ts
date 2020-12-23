import { PreventIframe } from "express-msteams-host";

/**
 * Used as place holder for the decorators
 */
@PreventIframe("/qnATab/index.html")
@PreventIframe("/qnATab/config.html")
@PreventIframe("/qnATab/remove.html")
export class QnATab {
}
