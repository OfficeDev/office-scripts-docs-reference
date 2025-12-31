export declare namespace OfficeScript {
    /**
     * Saves a copy of the current workbook in OneDrive, in the same directory as the original file, with the specified file name.
     * This API must be called before other APIs.
     * @param filename - The file name of the copied and saved file. The file name must end with ".xlsx".
     
     * **Throws**:  `SaveCopyAsInvalidExtension` Thrown if the file name doesn't end with ".xlsx".
     
     * **Throws**:  `SaveCopyAsMustBeCalledFirst` Thrown if this method is called after other APIs.
     
     * **Throws**:  `SaveCopyAsFileMayAlreadyExist` Thrown if the file name of the copy already exists.
     
     * **Throws**:  `SaveCopyAsInvalidCharacters` Thrown if the file name contains invalid characters.
     
     * **Throws**:  `SaveCopyAsFileNotOnOneDrive` Thrown if the document is not saved to OneDrive.
     
     * **Throws**:  `ExternalApiTimeout` Thrown if the API reaches the timeout limit of 30 seconds. Note that the copy may still be created.
     */
    export function saveCopyAs(filename: string): void;

    /**
     * Converts the document to a PDF and returns the text encoding of it.
     * @returns The content of the workbook as a string, in PDF format.
     
     * **Throws**:  `ConvertToPdfEmptyWorkbook` Thrown if the document is empty.
     
     * **Throws**:  `ConvertToPdfProtectedWorkbook` Thrown if the document is protected.
     
     * **Throws**:  `ExternalApiTimeout` Thrown if the API reaches the timeout limit of 30 seconds.
     */
    export function convertToPdf(): string;

    /**
     * Downloads a specified file to the default download location specified by the local machine.
     * @param fileProperties - The file to download.
     
     * **Throws**:  `DownloadFileNameMissing` Thrown if the name is empty.
     
     * **Throws**:  `DownloadFileContentMissing` Thrown if the content is empty.
     
     * **Throws**:  `DownloadFileInvalidExtension` Thrown if the file name extension is not ".txt" or ".pdf".
     
     * **Throws**:  `ExternalApiTimeout` Thrown if the API reaches the timeout limit of 30 seconds.
     */
    export function downloadFile(fileProperties: DownloadFileProperties): void;

    /**
     * The file to download.
     */
    export interface DownloadFileProperties {
        /**
         * The name of the file once downloaded. The file extension determines the type of the file. Supported extensions are ".txt" and ".pdf". Default is ".txt".
         */
        name: string;

        /**
         * The content of the file.
         */
        content: string;
    }

    /**
     * Send an email with an Office Script. Use `MailProperties` to specify the content and recipients of the email.
     * @param message - The properties that define the content and recipients of the email.
     
     * **Throws**:  `SendMailMaxCalls` Thrown if the maximum number of API calls is exceeded. The limit is 100 API calls.
     
     * **Throws**:  `SendMailNoRecipient` Thrown if no recipient is specified.
     
     * **Throws**:  `SendMailInvalidEmail` Thrown if an invalid email address is provided.
     
     * **Throws**:  `SendMailExtensionNotSupported` Thrown if the attachment name extension is not ".txt" or ".pdf".
     
     * **Throws**:  `ExternalApiTimeout` Thrown if the API reaches the timeout limit of 30 seconds.
     */
    export function sendMail(mailProperties: MailProperties): void;

    /**
     * The type of the content. Possible values are text or HTML.
     */
    enum EmailContentType {
        /**
         * The email message body is in HTML format.
         */
        html = "html",

        /**
         * The email message body is in plain text format.
         */
        text = "text",
    }

    /**
     * The importance value of the email. Corresponds to "high", "normal", and "low" importance values available in the Outlook UI.
     */
    enum EmailImportance {
        /**
         * Email is marked as low importance.
         */
        low = "low",

        /**
         * Email does not have any importance specified.
         */
        normal = "normal",

        /**
         * Email is marked as high importance.
         */
        high = "high",
    }

    /**
     * The attachment to send with the email.
     * A value must be specified for at least one of the `to`, `cc`, or `bcc` parameters.
     * If no recipient is specified, the following error is shown: "The message has no recipient. Please enter a value for at least one of the "to", "cc", or "bcc" parameters."
     */
    export interface EmailAttachment {
        /**
         * The text that is displayed below the icon representing the attachment. This string doesn't need to match the file name.
         */
        name: string;
        /**
         * The contents of the file.
         */
        content: string;
    }

    /**
     * The properties of the email to be sent.
     */
    export interface MailProperties {
        /**
         * The subject of the email. Optional.
         */
        subject?: string;

        /**
         * The content of the email. Optional.
         */
        content?: string;

        /**
         * The type of the content in the email. Possible values are text or HTML. Optional.
         */
        contentType?: EmailContentType;

        /**
         * The importance of the email. The possible values are `low`, `normal`, and `high`. Default value is `normal`. Optional.
         */
        importance?: EmailImportance;

        /**
         * The direct recipient or recipients of the email. Optional.
         */
        to?: string | string[];

        /**
         * The carbon copy (CC) recipient or recipients of the email. Optional.
         */
        cc?: string | string[];

        /**
         * The blind carbon copy (BCC) recipient or recipients of the email. Optional.
         */
        bcc?: string | string[];

        /**
         * A file (such as a text file or Excel workbook) attached to a message. Optional.
         */
        attachments?: EmailAttachment | EmailAttachment[];
    }

    /**
     * Metadata about the script.
     */
    export namespace Metadata {
        /**
         * Get the name of the currently running script.
         */
        export function getScriptName(): string;
    }
}