/* global Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        // Register the event handler for the ItemSend event
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, onItemSend);
        console.log('Add-in is running in background.');
    }
});

// async function fetchPolicyDomains() {
//     try {
//         const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Admin/policies', {
//             method: 'GET',
//             headers: {
//                 'Content-Type': 'application/json',
//                 'Accept': 'application/json'
//             }
//         });

//         if (!response.ok) {
//             throw new Error('Failed to fetch policy domains: ' + response.statusText);
//         }

//         const policies = await response.json();

//         const allowedDomains = policies[0].AllowedDomains || [];
//         const blockedDomains = policies[0].BlockedDomains || [];

//         return { allowedDomains, blockedDomains };
//     } catch (error) {
//         console.error('Error fetching policy domains:', error);
//         return { allowedDomains: [], blockedDomains: [] };
//     }
// }

// Event handler for the ItemSend event
async function onItemSend(eventArgs) {
    try {
        const item = Office.context.mailbox.item;

        // Get the necessary data
        const from = await getFromAsync(item);
        const toRecipients = await getRecipientsAsync(item.to);
        const ccRecipients = await getRecipientsAsync(item.cc);
        const bccRecipients = await getRecipientsAsync(item.bcc);
        const subject = await getSubjectAsync(item);
        const body = await getBodyAsync(item);
        const attachments = await getAttachmentsAsync(item);

         // Fetch the policy domains
        //  const { allowedDomains, blockedDomains } = await fetchPolicyDomains();

        //  // Check if the email's recipients are blocked or allowed
        //  if (isDomainBlockedOrAllowed(toRecipients, blockedDomains, allowedDomains) ||
        //      isDomainBlockedOrAllowed(ccRecipients, blockedDomains, allowedDomains) ||
        //      isDomainBlockedOrAllowed(bccRecipients, blockedDomains, allowedDomains)) {
 
        //      Office.context.mailbox.item.notificationMessages.addAsync("error", {
        //          type: "errorMessage",
        //          message: "This email cannot be sent as it contains blocked domains."
        //      });
        //      eventArgs.completed({ allowEvent: false });
        //      return;
        //  }
 
         // Prepare email data for saving
         const emailData = prepareEmailData(from, toRecipients, ccRecipients, bccRecipients, subject, body, attachments);
 
         // Save email data to the backend server
         await saveEmailData(emailData);
 
         // Allow the email to be sent
         eventArgs.completed();
     } catch (error) {
         console.error('Error during send event:', error);
         eventArgs.completed({ allowEvent: false });
     }
 }

// Helper function to prepare email data
function prepareEmailData(from, to, cc, bcc, subject, body, attachments) {
    let emailId = generateUUID();
    return {
        Id: emailId,
        FromEmailID: from,
        Attachments: attachments.map(attachment => ({
            Id: generateUUID(),
            FileName: attachment.name,
            FileType: attachment.attachmentType,
            FileSize: attachment.size,
            UploadTime: new Date().toISOString()
        })),
        EmailBcc: bcc ? bcc.split(',').map(email => email.trim()) : [],
        EmailCc: cc ? cc.split(',').map(email => email.trim()) : [],
        EmailBody: body,
        EmailSubject: subject,
        EmailTo: to ? to.split(',').map(email => email.trim()) : [],
        Timestamp: new Date().toISOString()
    };
}

// function isDomainBlockedOrAllowed(recipients, blockedDomains, allowedDomains) {
//     const recipientArray = recipients ? recipients.split(',').map(email => email.trim()) : [];
    
//     // Check if the recipient domain is blocked or allowed
//     for (let recipient of recipientArray) {
//         const domain = recipient.split('@')[1];  // Extract the domain from the email

//         if (blockedDomains.includes(domain)) {
//             console.log(`Domain ${domain} is blocked.`);
//             return true;
//         }

//         if (allowedDomains.length > 0 && !allowedDomains.includes(domain)) {
//             console.log(`Domain ${domain} is not in the allowed list.`);
//             return true;
//         }
//     }

//     return false;
// }


// Helper function to send the email data to the backend server
async function saveEmailData(emailData) {
    const response = await fetch('https://kntrolemail.kriptone.com:6677/api/Email', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
        },
        body: JSON.stringify(emailData)
    });

    if (!response.ok) {
        throw new Error('Failed to save email data: ' + response.statusText);
    }

    console.log('Email data saved successfully.');
}

// Helper function to generate a UUID
function generateUUID() { 
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        const r = Math.random() * 16 | 0, v = c === 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}

// Async functions to retrieve email details
function getFromAsync(item) {
    return new Promise((resolve, reject) => {
        item.from.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const fromEmail = result.value.emailAddress || result.value.address;
                resolve(fromEmail);
            } else {
                reject('Error retrieving from address: ' + result.error.message);
            }
        });
    });
}

function getRecipientsAsync(recipients) {
    return new Promise((resolve, reject) => {
        recipients.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const emails = result.value.map(recipient => recipient.emailAddress || recipient.address).join(", ");
                resolve(emails);
            } else {
                reject('Error retrieving recipients: ' + result.error.message);
            }
        });
    });
}

function getSubjectAsync(item) {
    return new Promise((resolve, reject) => {
        item.subject.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject('Error retrieving subject: ' + result.error.message);
            }
        });
    });
}

function getBodyAsync(item) {
    return new Promise((resolve, reject) => {
        item.body.getAsync(Office.CoercionType.Text, function(result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject('Error retrieving body: ' + result.error.message);
            }
        });
    });
}

function getAttachmentsAsync(item) {
    return new Promise((resolve, reject) => {
        item.getAttachmentsAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const attachments = result.value;
                resolve(attachments || []);
            } else {
                reject('Error retrieving attachments: ' + result.error.message);
            }
        });
    });
}
