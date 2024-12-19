Office.initialize = function (reason) {
    // Initialize the Office.js library.
    console.log("Office initialized with reason: " + reason);
};

// Function to insert the signature
function insertSignature(event) {
    // Define the signature content
    const signature = `
        <br><br>
        --<br>
        <strong>Your Name</strong><br>
        Your Position<br>
        <a href="mailto:your.email@example.com">your.email@example.com</a><br>
        <a href="https://www.yourcompany.com">www.yourcompany.com</a>
    `;

    // Use Office.js to set the body of the email
    Office.context.mailbox.item.body.setAsync(
        signature,
        { coercionType: Office.CoercionType.Html }, // Ensure the content is set as HTML
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.error("Failed to insert signature: " + asyncResult.error.message);
            } else {
                console.log("Signature inserted successfully.");
            }
        }
    );

    // Signal that the event handling is complete
    event.completed();
}