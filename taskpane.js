/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
        console.log("✅ Add-in is running in the background.");
    }
});

// Email regex: validates general email format with 2-3 character domain extensions
const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,3}$/;

// Regex patterns for additional checks
const regexPatterns = {
    email: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,3}$/, 
    body: /\b(confidential|prohibited|restricted)\b/i,
    attachmentName: /\.(exe|bat|sh)$/i,
    imei: /[^\d](\d{15}|\d{2}\-\d{6}\-\d{6}\-\d)[^\d]/,
    namesUSCensus: /[^\w]([A-Z][a-z]{1,12}(\s?,\s?|[\s]|\s([A-Z])\.\s)[A-Z][a-z]{1,12})[^\w]/,
    swiftBIC: /[^\w-]([A-Z]{6}[A-Z0-9]{2}([A-Z0-9]{3})?)[^\w-]/,
    timeZoneOffset: /[^\d]([-+](0[0-9]|1[0-2]):00|\+13:00|[+-]0[34]:30|\+0(5:(30|45)|9:30))[^\d]/,
    ibanAndorra: /[^\w](AD\d{2}((\s\d{4}){2}(\s[a-zA-Z0-9]){3}|\d{8}[a-zA-Z0-9]{12}))[^\w]/,
    taxFileNumber: /[^\w-.;&](\d{8,9})[^\w-.;&]/,
    ibanAustria: /[^\w](AT\d{2}(\s(\d{4}\s){3}\d{4}|\d{16}))[^\w]/,
    phoneAustria: /[^\d\-]((\+43[\s\-]|0)(\d{1,4}[\s\-]\d{3}[\s\-]\d{3}([\s\-]\d{3,6})?|\d{1,4}[\s\-]\d{3,12}))[^\d\-]/,
    ssnAustria: /[^\w-.;&](\d{4}[0-3]\d(0[1-9]|1[0-2])\d{2})[^\w-.;&]/,
    ibanAzerbaijan: /[^\w](AZ\d{2}(\s[A-Za-z0-9]{4}\s(\d{4}\s){4}\d{4}|[A-Za-z0-9]{4}\d{20}))[^\w]/,
};

async function onMessageSendHandler(eventArgs) {
    try {
        const item = Office.context.mailbox.item;
        const from = await getFromAsync(item);
        const toRecipients = await getRecipientsAsync(item.to);
        const ccRecipients = await getRecipientsAsync(item.cc);
        const bccRecipients = await getRecipientsAsync(item.bcc);
        const subject = await getSubjectAsync(item);
        const body = await getBodyAsync(item);
        const attachments = await getAttachmentsAsync(item);

        console.log("🔹 Email Details:", { from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments });

        const { allowedDomains, blockedDomains } = await fetchPolicyDomains();
        console.log("🔹 Policy Check:", { allowedDomains, blockedDomains });

        if (isDomainBlocked(toRecipients, blockedDomains) || 
            isDomainBlocked(ccRecipients, blockedDomains) || 
            isDomainBlocked(bccRecipients, blockedDomains)) {
            showOutlookNotification("Blocked Domain", "Your email contains a blocked domain. Sending is prevented.");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        if (!toRecipients && !ccRecipients && !bccRecipients) {
            showOutlookNotification("Missing Recipients", "Please add at least one recipient (To, CC, or BCC).");
            eventArgs.completed({ allowEvent: false });
            return;
        }

        for (const pattern in regexPatterns) {
            if (regexPatterns[pattern].test(body)) {
                showOutlookNotification("Restricted Content", `Your email contains restricted content: ${pattern}.`);
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        for (const attachment of attachments) {
            if (regexPatterns.attachmentName.test(attachment.name)) {
                showOutlookNotification("Restricted Attachment", `Attachment \"${attachment.name}\" is restricted.`);
                eventArgs.completed({ allowEvent: false });
                return;
            }
        }

        console.log("✅ Passed all policy checks. Saving email data...");
        const emailData = prepareEmailData(from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments);
        const saveSuccess = await saveEmailData(emailData);

        if (saveSuccess.success) {
            console.log("✅ Email data saved. Ensuring email is sent.");
            eventArgs.completed({ allowEvent: true });
        } else {
            showOutlookNotification("Error", saveSuccess.message || "Failed to save email data.");
            eventArgs.completed({ allowEvent: false });
        }
    } catch (error) {
        console.error('❌ Error during send event:', error);
        showOutlookNotification("Error", "An unexpected error occurred.");
        eventArgs.completed({ allowEvent: false });
    }
}

async function fetchPolicyDomains() {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Admin/policies', {
            method: 'GET',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
        });

        if (!response.ok) throw new Error('Failed to fetch policy domains.');

        const json = await response.json();
        return { 
            allowedDomains: json.data[0]?.allowedDomains || [], 
            blockedDomains: json.data[0]?.blockedDomains || [] 
        };
    } catch (error) {
        console.error("❌ Error fetching policy domains:", error);
        return { allowedDomains: [], blockedDomains: [] };
    }
}

async function saveEmailData(emailData) {
    try {
        const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
            body: JSON.stringify(emailData),
        });
        
        const json = await response.json();
        return { success: response.ok && json.success, message: json.message || "Unknown error" };
    } catch (error) {
        console.error("❌ Error saving email data:", error);
        return { success: false, message: "Unable to connect to the server." };
    }
}

function isDomainBlocked(recipients, blockedDomains) {
    if (!blockedDomains || blockedDomains.length === 0) return false;
    const recipientArray = recipients ? recipients.split(',').map(email => email.trim()) : [];
    return recipientArray.some(email => blockedDomains.includes(email.split('@')[1]));
}

function prepareEmailData(from, to, cc, bcc, subject, body, attachments) {
    return {
        Id: generateUUID(),
        FromEmailID: from,
        Attachments: attachments.map(attachment => ({
            Id: generateUUID(),
            FileName: attachment.name,
            FileType: attachment.attachmentType,
            FileSize: attachment.size,
            UploadTime: new Date().toISOString(),
        })),
        EmailBcc: bcc ? bcc.split(',').map(email => email.trim()) : [],
        EmailCc: cc ? cc.split(',').map(email => email.trim()) : [],
        EmailBody: body,
        EmailSubject: subject,
        EmailTo: to ? to.split(',').map(email => email.trim()) : [],
        Timestamp: new Date().toISOString(),
    };
}


function validateEmailAddresses(recipients) {
    return recipients ? recipients.split(',').every(email => emailRegex.test(email.trim())) : true;
}
function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
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

function showOutlookNotification(title, message) {
    Office.context.mailbox.item.notificationMessages.addAsync("error", {
        type: "errorMessage",
        message: `${title}: ${message}`,
    });
}