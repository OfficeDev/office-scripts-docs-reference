export declare namespace OfficeScript {
    /**
     * Saves a copy of the current workbook in OneDrive, in the same directory as the original file, with the specified file name.
     * The API has a timeout limit of 30 seconds. This limit is rarely exceeded.
     * Note: Timeout doesn't necessarily indicate that the API failed. The workbook copy may still be created, but after the timeout limit this API does not return a success or failure message.
     * @throws ExternalApiTimeout The error thrown if the API reaches the timeout limit of 30 seconds. Note that the copy may still be created.
     * @throws InvalidExtensionError The error thrown if the file name doesn't end with ".xlsx".
     * @throws SaveCopyAsFileMayAlreadyExistError The error thrown if the file name of the copy already exists.
     * @throws SaveCopyAsFileNotOnOneDriveError The error thrown if the document is not saved to OneDrive.
     * @param filename - The file name of the copied and saved file. The file name must end with ".xlsx".
     * @beta
     */
    export function saveCopyAs(filename: string): void;

    /**
     * Return the text encoding of the document as a PDF.
     * If the document is empty, then the following error is shown: "We didn't find anything to print".
     * Some actions made prior to using this API may not be captured in the PDF in Excel on the web.
     * @returns The content of the workbook as a string, in PDF format.
     * @beta
     */
    export function convertToPdf(): string;

    /**
     * Downloads a specified file to the default download location specified by the local machine.
     * @param name - The name of the file once downloaded. The file extension determines the type of the file. Supported extensions are ".txt" and ".pdf". Default is ".txt".
     * @param content - The content of the file.
     * @beta
     */
    export function downloadFile({
        name,
        content,
    }: {
        name: string;
        content: string;
    }): void;

    /**
     * Send an email with an Office Script. Use `MailProperties` to specify the content and recipients of the email.
     * If the request body includes content, this method returns 400 Bad request.
     * @param message - The properties that define the content and recipients of the email.
     * @beta
     */
    export function sendMail(mailProperties: MailProperties): void;

    /**
     * The type of the content. Possible values are text or HTML.
     * @beta
     */
    enum EmailContentType {
        /**
         * The email message body is in HTML format.
         * @beta
         */
        html = "html",

        /**
         * The email message body is in plain text format.
         * @beta
         */
        text = "text",
    }

    /**
     * The importance value of the email. Corresponds to "high", "normal", and "low" importance values available in the Outlook UI.
     * @beta
     */
    enum EmailImportance {
        /**
         * Email is marked as low importance.
         * @beta
         */
        low = "low",

        /**
         * Email does not have any importance specified.
         * @beta
         */
        normal = "normal",

        /**
         * Email is marked as high importance.
         * @beta
         */
        high = "high",
    }

    /**
     * The attachment to send with the email.
     * A value must be specified for at least one of the `to`, `cc`, or `bcc` parameters.
     * If no recipient is specified, the following error is shown: "The message has no recipient. Please enter a value for at least one of the "to", "cc", or "bcc" parameters."
     * @beta
     */
    export interface EmailAttachment {
        /**
         * The text that is displayed below the icon representing the attachment. This string doesn't need to match the file name.
         * @beta
         */
        name: string;
        /**
         * The contents of the file.
         * @beta
         */
        content: string;
    }

    /**
     * The properties of the email to be sent.
     * @beta
     */
    export interface MailProperties {
        /**
         * The subject of the email. Optional.
         * @beta
         */
        subject?: string;

        /**
         * The content of the email. Optional.
         * @beta
         */
        content?: string;

        /**
         * The type of the content in the email. Possible values are text or HTML. Optional.
         * @beta
         */
        contentType?: EmailContentType;

        /**
         * The importance of the email. The possible values are `low`, `normal`, and `high`. Default value is `normal`. Optional.
         * @beta
         */
        importance?: EmailImportance;

        /**
         * The direct recipient or recipients of the email. Optional.
         * @beta
         */
        to?: string | string[];

        /**
         * The carbon copy (CC) recipient or recipients of the email. Optional.
         * @beta
         */
        cc?: string | string[];

        /**
         * The blind carbon copy (BCC) recipient or recipients of the email. Optional.
         * @beta
         */
        bcc?: string | string[];

        /**
         * A file (such as a text file or Excel workbook) attached to a message. Optional.
         * @beta
         */
        attachments?: EmailAttachment | EmailAttachment[];
    }

    /**
     * Metadata about the script.
     * @beta
     */
    export namespace Metadata {
        /**
         * Get the name of the currently running script.
         * @beta
         */
        export function getScriptName(): string;
    }
}