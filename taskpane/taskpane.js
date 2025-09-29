Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    // Office is ready
  }
});

function modifySubjectAndForward(letter) {
  Office.context.mailbox.item.subject.getAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const originalSubject = result.value;
      const newSubject = originalSubject + ' ' + letter;

      Office.context.mailbox.item.subject.setAsync(newSubject, function(setResult) {
        if (setResult.status === Office.AsyncResultStatus.Succeeded) {
          Office.context.mailbox.item.forwardAsync({
            toRecipients: ['filip.skalka@zpmvcr.cz'],
            subject: newSubject
          }, { htmlBody: 'Forwarded with subject modification.' }, function(forwardResult) {
            if (forwardResult.status === Office.AsyncResultStatus.Succeeded) {
              sendConfirmationEmail(newSubject);
            }
          });
        }
      });
    }
  });
}

function sendConfirmationEmail(subject) {
  const userEmail = Office.context.mailbox.userProfile.emailAddress;
  const item = {
    toRecipients: [userEmail],
    subject: 'Confirmation: Email forwarded with subject "' + subject + '"',
    htmlBody: '<p>Your email was forwarded successfully.</p>'
  };
  Office.context.mailbox.item.notificationMessages.addAsync('confirmation', {
    type: 'informationalMessage',
    message: 'Confirmation email sent.',
    icon: 'icon16',
    persistent: false
  });
  Office.context.mailbox.item.saveAsync(function() {
    Office.context.mailbox.item.displayReplyForm(item.htmlBody);
  });
}
