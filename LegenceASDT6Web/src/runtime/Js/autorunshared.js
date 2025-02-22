// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

// Contains code for event-based activation on Outlook on web, on Windows, and on Mac (new UI preview).

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * @param {*} eventObj Office event object
 * @returns
 */

//This is a test
async function loadSignatureFromFile() {
    // Append a timestamp to force a fresh fetch
    const filePath = `https://siggy.wearelegence.com/users/corey.gashlin@wearelegence.com.html?cb=${new Date().getTime()}`;
    try {
        const response = await fetch(filePath, { cache: "no-store" }); // no-store ensures no caching
        if (!response.ok) {
            throw new Error(`Failed to load file: ${response.status} ${response.statusText}`);
        }
        return await response.text();
    } catch (error) {
        console.error('Error fetching HTML file:', error);
        return null;
    }
}

async function checkSignature(eventObj) {
    const signature = await loadSignatureFromFile();

    if (signature) {
        Office.context.mailbox.item.body.setSignatureAsync(
            signature,
            { coercionType: "html" },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Signature applied successfully.");
                } else {
                    console.error('Failed to set signature:', asyncResult.error.message);
                }
                eventObj.completed();
            }
        );
    } else {
        console.error('No signature loaded.');
        eventObj.completed();
    }
}

/**
 * For Outlook on Windows and on Mac only. Insert signature into appointment or message.
 * Outlook on Windows and on Mac can use setSignatureAsync method on appointments and messages.
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @param {*} user_info Information details about the user
 * @param {*} eventObj Office event object
 */
function insert_auto_signature(compose_type, user_info, eventObj) {
    let template_name = get_template_name(compose_type);
    let signature_info = get_signature_info(template_name, user_info);
    addTemplateSignature(signature_info, eventObj);
}

/**
 * 
 * @param {*} signatureDetails object containing:
 *  "signature": The signature HTML of the template,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 * @param {*} eventObj 
 * @param {*} signatureImageBase64 
 */
function addTemplateSignature(signatureDetails, eventObj, signatureImageBase64) {
    if (is_valid_data(signatureDetails.logoBase64) === true) {
        //If a base64 image was passed we need to attach it.
        Office.context.mailbox.item.addFileAttachmentFromBase64Async(
            signatureDetails.logoBase64,
            signatureDetails.logoFileName,
            {
                isInline: true,
            },
            function (result) {
                //After image is attached, insert the signature
                Office.context.mailbox.item.body.setSignatureAsync(
                    signatureDetails.signature,
                    {
                        coercionType: "html",
                        asyncContext: eventObj,
                    },
                    function (asyncResult) {
                        asyncResult.asyncContext.completed();
                    }
                );
            }
        );
    } else {
        //Image is not embedded, or is referenced from template HTML
        Office.context.mailbox.item.body.setSignatureAsync(
            signatureDetails.signature,
            {
                coercionType: "html",
                asyncContext: eventObj,
            },
            function (asyncResult) {
                asyncResult.asyncContext.completed();
            }
        );
    }
}

/**
 * Creates information bar to display when new message or appointment is created
 */
function display_insight_infobar() {
    Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a", {
        type: "insightMessage",
        message: "Please set your signature with the Office Add-ins sample.",
        icon: "Icon.16x16",
        actions: [
            {
                actionType: "showTaskPane",
                actionText: "Set signatures",
                commandId: get_command_id(),
                contextData: "{''}",
            },
        ],
    });
}

/**
 * Gets template name (A,B,C) mapped based on the compose type
 * @param {*} compose_type The compose type (reply, forward, newMail)
 * @returns Name of the template to use for the compose type
 */
function get_template_name(compose_type) {
    if (compose_type === "reply") return Office.context.roamingSettings.get("reply");
    if (compose_type === "forward") return Office.context.roamingSettings.get("forward");
    return Office.context.roamingSettings.get("newMail");
}

/**
 * Gets HTML signature in requested template format for given user
 * @param {\} template_name Which template format to use (A,B,C)
 * @param {*} user_info Information details about the user
 * @returns HTML signature in requested template format
 */
function get_signature_info(template_name, user_info) {
    if (template_name === "templateB") return get_template_B_info(user_info);
    if (template_name === "templateC") return get_template_C_info(user_info);
    return get_template_A_info(user_info);
}

/**
 * Gets correct command id to match to item type (appointment or message)
 * @returns The command id
 */
function get_command_id() {
    if (Office.context.mailbox.item.itemType == "appointment") {
        return "MRCS_TpBtn1";
    }
    return "MRCS_TpBtn0";
}

/**
 * Gets HTML string for template A
 * Embeds the signature logo image into the HTML string
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template A,
    "logoBase64": The base64 encoded logo image,
    "logoFileName": The filename of the logo image
 */
function get_template_A_info(user_info) {
    const logoFileName = "sample-logo.png";
    let str = "";
    if (is_valid_data(user_info.greeting)) {
        str += user_info.greeting + "<br/>";
    }

    str += "<table>";
    str += "<tr>";
    // Embed the logo using <img src='cid:...
    str +=
        "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:" +
        logoFileName +
        "' alt='MS Logo' width='24' height='24' /></td>";
    str += "<td style='padding-left: 5px;'>";
    str += "<strong>" + user_info.name + "</strong>";
    str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
    str += "<br/>";
    str += is_valid_data(user_info.job) ? user_info.job + "<br/>" : "";
    str += user_info.email + "<br/>";
    str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
    str += "</td>";
    str += "</tr>";
    str += "</table>";

    // return object with signature HTML, logo image base64 string, and filename to reference it with.
    return {
        signature: str,
        logoBase64:
            "iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",
        logoFileName: logoFileName,
    };
}

/**
 * Gets HTML string for template B
 * References the signature logo image from the HTML
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template B,
    "logoBase64": null since this template references the image and does not embed it ,
    "logoFileName": null since this template references the image and does not embed it
 */
function get_template_B_info(user_info) {
    let str = "";
    if (is_valid_data(user_info.greeting)) {
        str += user_info.greeting + "<br/>";
    }

    str += "<table>";
    str += "<tr>";
    // Reference the logo using a URI to the web server <img src='https://...
    str +=
        "<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>";
    str += "<td style='padding-left: 5px;'>";
    str += "<strong>" + user_info.name + "</strong>";
    str += is_valid_data(user_info.pronoun) ? "&nbsp;" + user_info.pronoun : "";
    str += "<br/>";
    str += user_info.email + "<br/>";
    str += is_valid_data(user_info.phone) ? user_info.phone + "<br/>" : "";
    str += "</td>";
    str += "</tr>";
    str += "</table>";

    return {
        signature: str,
        logoBase64: null,
        logoFileName: null,
    };
}

/**
 * Gets HTML string for template C
 * @param {*} user_info Information details about the user
 * @returns Object containing:
 *  "signature": The signature HTML of template C,
    "logoBase64": null since there is no image,
    "logoFileName": null since there is no image
 */
function get_template_C_info(user_info) {
    let str = "";
    if (is_valid_data(user_info.greeting)) {
        str += user_info.greeting + "<br/>";
    }

    str += user_info.name;

    return {
        signature: str,
        logoBase64: null,
        logoFileName: null,
    };
}

/**
 * Validates if str parameter contains text.
 * @param {*} str String to validate
 * @returns true if string is valid; otherwise, false.
 */
function is_valid_data(str) {
    return str !== null && str !== undefined && str !== "";
}

Office.actions.associate("checkSignature", checkSignature);