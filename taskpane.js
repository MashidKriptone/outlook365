/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Register the event handler for the ItemSend event
        Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
        console.log('Add-in is running in the background.');
    }
});

// Email regex: validates general email format with 2-3 character domain extensions
const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,3}$/;

// Regex patterns for additional checks
const regexPatterns = {
    body: /\b(confidential|prohibited|restricted)\b/i, // Example sensitive keywords in the body
    attachmentName: /\.(exe|bat|sh)$/i, // Example restricted file extensions
};

// Event handler for the ItemSend event
async function onMessageSendHandler(eventArgs) {
    try {
        const item = Office.context.mailbox.item;

        // Retrieve email details
        const from = await getFromAsync(item);
        const toRecipients = await getRecipientsAsync(item.to);
        const ccRecipients = await getRecipientsAsync(item.cc);
        const bccRecipients = await getRecipientsAsync(item.bcc);
        const subject = await getSubjectAsync(item);
        const body = await getBodyAsync(item);
        const attachments = await getAttachmentsAsync(item);

        console.log("🔹 Email Details:");
        console.log("From:", from);
        console.log("To Recipients:", toRecipients);
        console.log("CC Recipients:", ccRecipients);
        console.log("BCC Recipients:", bccRecipients);
        console.log("Subject:", subject);
        console.log("Body:", body);
        console.log("Attachments:", attachments);

        // Fetch policy domains
        const { allowedDomains, blockedDomains } = await fetchPolicyDomains();

        console.log("🔹 Policy Check:");
        console.log("Allowed Domains:", allowedDomains);
        console.log("Blocked Domains:", blockedDomains);

        // Allow email if no policies are defined
        if (allowedDomains.length === 0 && blockedDomains.length === 0) {
            console.log("✅ No policy restrictions found. Email will be sent.");
            eventArgs.completed(); // Ensure email is sent
            return;
        }

        // Check if email is blocked due to blocked domains
        if (
            blockedDomains.length > 0 &&
            (isDomainBlocked(toRecipients, blockedDomains) ||
                isDomainBlocked(ccRecipients, blockedDomains) ||
                isDomainBlocked(bccRecipients, blockedDomains))
        ) {
            console.warn("❌ Blocked domain detected. Email is not sent.");
            Office.context.mailbox.item.notificationMessages.addAsync("error", {
                type: "errorMessage",
                message: "KntrolEMAIL detected a blocked domain policy and prevented the email from being sent.",
            });
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Validate email addresses
        if (!validateEmailAddresses(toRecipients) ||
            !validateEmailAddresses(ccRecipients) ||
            !validateEmailAddresses(bccRecipients)) {
            console.warn("❌ Invalid email addresses found. Email is not sent.");
            Office.context.mailbox.item.notificationMessages.addAsync("error", {
                type: "errorMessage",
                message: "One or more email addresses are invalid.",
            });
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Validate body content
        if (regexPatterns.body.test(body)) {
            console.warn("❌ Prohibited content detected in the email body. Email is not sent.");
            Office.context.mailbox.item.notificationMessages.addAsync("error", {
                type: "errorMessage",
                message: "The email contains prohibited content in the body.",
            });
            eventArgs.completed({ allowEvent: false });
            return;
        }

        // Validate attachments
        for (const attachment of attachments) {
            if (regexPatterns.attachmentName.test(attachment.name)) {
                console.warn(`❌ Attachment ${attachment.name} is restricted. Email is not sent.`);
                Office.context.mailbox.item.notificationMessages.addAsync("error", {
                    type: "errorMessage",
                    message: `The attachment \"${attachment.name}\" has a restricted file type.`,
                });
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        console.log("✅ All checks passed. Proceeding to save email data...");

        // Prepare email data for saving
        const emailData = prepareEmailData(from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments);

        // Save email data to the backend
        await saveEmailData(emailData);

        console.log("✅ Email data saved successfully. Email will be sent.");
        
        // Allow the email to be sent
        eventArgs.completed();
    } catch (error) {
        console.error('❌ Error during send event:', error);
        eventArgs.completed({ allowEvent: false });
    }
}

// Fetch policy domains from the backend
async function fetchPolicyDomains() {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Admin/policies', {
            method: 'GET',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
            },
        });

        if (!response.ok) {
            throw new Error('Failed to fetch policy domains: ' + response.statusText);
        }

        const policies = await response.json();
        return { 
            allowedDomains: policies.data[0]?.allowedDomains || [], 
            blockedDomains: policies.data[0]?.blockedDomains || [] 
        };
    } catch (error) {
        console.error('❌ Error fetching policy domains:', error);
        return { allowedDomains: [], blockedDomains: [] };
    }
}

// Save email data to the backend
async function saveEmailData(emailData) {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Accept': 'application/json',
            },
            body: JSON.stringify(emailData),
        });

        if (!response.ok) {
            throw new Error('Failed to save email data: ' + response.statusText);
        }

        console.log('✅ Email data saved successfully.');
    } catch (error) {
        console.error('❌ Error saving email data:', error);
    }
}

// Helper function to validate email addresses
function validateEmailAddresses(recipients) {
    if (!recipients) return true;

    const emailArray = recipients.split(',').map(email => email.trim());
    for (let email of emailArray) {
        if (!emailRegex.test(email)) {
            console.log(`❌ Invalid email address: ${email}`);
            return false;
        }
    }
    return true;
}

// Helper function to check blocked domains
function isDomainBlocked(recipients, blockedDomains) {
    if (!blockedDomains || blockedDomains.length === 0) return false;

    const recipientArray = recipients ? recipients.split(',').map(email => email.trim()) : [];
    return recipientArray.some(email => blockedDomains.includes(email.split('@')[1]));
}

// Async functions to retrieve email details
function getFromAsync(item) {
    return new Promise((resolve, reject) => {
        item.from.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value.emailAddress) : reject(result.error));
    });
}

function getRecipientsAsync(recipients) {
    return new Promise((resolve, reject) => {
        recipients.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value.map(r => r.emailAddress).join(", ")) : reject(result.error));
    });
}

function getSubjectAsync(item) {
    return new Promise((resolve, reject) => {
        item.subject.getAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error));
    });
}

function getBodyAsync(item) {
    return new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error));
    });
}

function getAttachmentsAsync(item) {
    return new Promise((resolve, reject) => {
        item.getAttachmentsAsync(result => result.status === Office.AsyncResultStatus.Succeeded ? resolve(result.value) : reject(result.error));
    });
}
