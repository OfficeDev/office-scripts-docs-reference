export declare namespace OfficeScript {
    /**
     * Saves a copy of the current workbook in OneDrive, in the same directory as the original file, with the specified file name.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     * @param filename - The file name of the copied and saved file. The file name must end with ".xlsx".
     
     * **Throws**:  `InvalidExtensionError` The error thrown if the file name doesn't end with ".xlsx".
     
     * **Throws**:  `SaveCopyAsFileMayAlreadyExistError` The error thrown if the file name of the copy already exists.
     
     * **Throws**:  `SaveCopyAsErrorInvalidCharacters` The error thrown if the file name contains invalid characters.
     
     * **Throws**:  `SaveCopyAsFileNotOnOneDriveError` The error thrown if the document is not saved to OneDrive.
     
     * **Throws**:  `ExternalApiTimeout` The error thrown if the API reaches the timeout limit of 30 seconds. Note that the copy may still be created.
     */
    export function saveCopyAs(filename: string): void;

    /**
     * Converts the document to a PDF and returns the text encoding of it.
     * Note: Recent changes made to the workbook in Excel on the web, through Office Scripts or the Excel UI, may not be captured in the PDF.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     * @returns The content of the workbook as a string, in PDF format.
     
     * **Throws**:  `ConvertToPdfEmptyWorkbook` The error thrown if the document is empty.
     
     * **Throws**:  `ConvertToPdfProtectedWorkbook` The error thrown if the document is protected.
     
     * **Throws**:  `ExternalApiTimeout` The error thrown if the API reaches the timeout limit of 30 seconds.
     */
    export function convertToPdf(): string;

    /**
     * Downloads a specified file to the default download location specified by the local machine.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     * @param fileProperties - The file to download.
     
     * **Throws**:  `DownloadFileNameMissing` The error thrown if the name is empty.
     
     * **Throws**:  `DownloadFileContentMissing` The error thrown if the content is empty.
     
     * **Throws**:  `DownloadFileInvalidExtension` The error thrown if the file name extension is not ".txt" or ".pdf".
     
     * **Throws**:  `ExternalApiTimeout` The error thrown if the API reaches the timeout limit of 30 seconds.
     */
    export function downloadFile(fileProperties: FileProperties): void;

    /**
     * The file to download.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     */
    export interface FileProperties {
        /**
         * The name of the file once downloaded. The file extension determines the type of the file. Supported extensions are ".txt" and ".pdf". Default is ".txt".
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        name: string;

        /**
         * The content of the file.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        content: string;
    }

    /**
     * Send an email with an Office Script. Use `MailProperties` to specify the content and recipients of the email.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     * @param message - The properties that define the content and recipients of the email.
     
     * **Throws**:  `SendMailErrorMaxCalls` The error thrown if the maximum number of API calls is exceeded. The limit is 100 API calls.
     
     * **Throws**:  `SendMailNoRecipient` The error thrown if no recipient is specified.
     
     * **Throws**:  `SendMailInvalidEmail` The error thrown if an invalid email address is provided.
     
     * **Throws**:  `SendMailExtensionNotSupported` The error thrown if the attachment name extension is not ".txt" or ".pdf".
     
     * **Throws**:  `ExternalApiTimeout` The error thrown if the API reaches the timeout limit of 30 seconds.
     */
    export function sendMail(mailProperties: MailProperties): void;

    /**
     * The type of the content. Possible values are text or HTML.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     */
    enum EmailContentType {
        /**
         * The email message body is in HTML format.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        html = "html",

        /**
         * The email message body is in plain text format.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        text = "text",
    }

    /**
     * The importance value of the email. Corresponds to "high", "normal", and "low" importance values available in the Outlook UI.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     */
    enum EmailImportance {
        /**
         * Email is marked as low importance.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        low = "low",

        /**
         * Email does not have any importance specified.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        normal = "normal",

        /**
         * Email is marked as high importance.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        high = "high",
    }

    /**
     * The attachment to send with the email.
     * A value must be specified for at least one of the `to`, `cc`, or `bcc` parameters.
     * If no recipient is specified, the following error is shown: "The message has no recipient. Please enter a value for at least one of the "to", "cc", or "bcc" parameters."
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     */
    export interface EmailAttachment {
        /**
         * The text that is displayed below the icon representing the attachment. This string doesn't need to match the file name.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        name: string;
        /**
         * The contents of the file.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        content: string;
    }

    /**
     * The properties of the email to be sent.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     */
    export interface MailProperties {
        /**
         * The subject of the email. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        subject?: string;

        /**
         * The content of the email. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        content?: string;

        /**
         * The type of the content in the email. Possible values are text or HTML. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        contentType?: EmailContentType;

        /**
         * The importance of the email. The possible values are `low`, `normal`, and `high`. Default value is `normal`. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        importance?: EmailImportance;

        /**
         * The direct recipient or recipients of the email. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        to?: string | string[];

        /**
         * The carbon copy (CC) recipient or recipients of the email. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        cc?: string | string[];

        /**
         * The blind carbon copy (BCC) recipient or recipients of the email. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        bcc?: string | string[];

        /**
         * A file (such as a text file or Excel workbook) attached to a message. Optional.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        attachments?: EmailAttachment | EmailAttachment[];
    }

    /**
     * Metadata about the script.
     * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
     */
    export namespace Metadata {
        /**
         * Get the name of the currently running script.
         * @beta This API is in preview and may change based on feedback. Do not use this API in a production environment.
         */
        export function getScriptName(): string;
    }
}