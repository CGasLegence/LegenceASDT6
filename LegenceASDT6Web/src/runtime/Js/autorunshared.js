// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

/**
 * Checks if signature exists.
 * If not, displays the information bar for user.
 * If exists, insert signature into new item (appointment or message).
 * Runs specific logic for Android/iOS platforms to clean HTML.
 * @param {*} eventObj Office event object
 * @returns
 */

async function loadSignatureFromFile() {
    const userEmail = Office.context.mailbox.userProfile.emailAddress;

    // Encode the email to ensure it's URL-safe
    const encodedEmail = encodeURIComponent(userEmail);
    const filePath = `https://siggy.wearelegence.com/users/${encodedEmail}.html?cb=${new Date().getTime()}`;
    try {
        const response = await fetch(filePath, { cache: "no-store" });
        if (!response.ok) {
            throw new Error(`Failed to load file: ${response.status} ${response.statusText}`);
        }
        return await response.text(); // Raw HTML content
    } catch (error) {
        console.error("Error fetching HTML file:", error);
        return null;
    }
}

function cleanHtmlForWhitespace(html) {
    // Create a hidden container to process the HTML
    const container = document.createElement("div");
    container.style.visibility = "hidden";
    container.innerHTML = html;
    document.body.appendChild(container);

    // Normalize margins and padding for all elements
    const allElements = container.querySelectorAll("*");
    allElements.forEach((el) => {
        el.style.margin = "0";
        el.style.padding = "0";
        el.style.lineHeight = "1.2"; // Adjust as needed for consistent spacing
    });

    // Extract cleaned-up HTML
    const cleanedHtml = container.innerHTML;
    document.body.removeChild(container);

    return cleanedHtml;
}

async function checkSignature(eventObj) {
    const platform = Office.context.mailbox.diagnostics.hostName.toLowerCase();


        console.log("Running logic for non-mobile platforms...");
        const signature = await loadSignatureFromFile();

        if (signature) {
            Office.context.mailbox.item.body.setSignatureAsync(
                signature,
                { coercionType: "html" },
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        console.log("Signature applied successfully.");
                    } else {
                        console.error("Failed to set signature:", asyncResult.error.message);
                    }
                    eventObj.completed();

                    const notification = {
                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                        message: "Desktop Signature added successfully",
                        icon: "none",
                        persistent: false,
                    };
                    Office.context.mailbox.item.notificationMessages.replaceAsync("signatureNotification", notification);
                }
            );
        } else {
            console.error("No signature loaded.");
            eventObj.completed();
        }
    
}

/**
 * Additional helper functions for templates or attachment-based logic (unchanged)
 * Insert auto signature, display insight info bar, or handle templates as needed
 * These remain as they are in your code.
 */

// Associate the action with the add-in
Office.actions.associate("checkSignature", checkSignature);
