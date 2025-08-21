/* global Office */

Office.onReady(() => {
  // Ready to use
});

async function getBodyText() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value || "");
        } else {
          reject(new Error(result.error?.message || "getAsync failed"));
        }
      }
    );
  });
}

async function setBodyText(text) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message || "setAsync failed"));
        }
      }
    );
  });
}

function notify(title, message) {
  try {
    Office.context.mailbox.item.notificationMessages.replaceAsync("status", {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: `${title}: ${message}`,
      icon: "icon",
      persistent: false
    });
  } catch (e) {
    // ignore
  }
}

async function callRewriteApi(text) {
  const res = await fetch("/api/rewrite", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ text })
  });
  const data = await res.json();
  if (!res.ok) {
    throw new Error(data.error || "rewrite failed");
  }
  return data.result;
}

// Exported command function
async function rewriteToBusiness(event) {
  try {
    notify("リライト", "処理を開始します");
    const original = await getBodyText();
    const rewritten = await callRewriteApi(original);
    await setBodyText(rewritten);
    notify("リライト", "完了しました");
  } catch (err) {
    notify("エラー", err.message || String(err));
  } finally {
    // Signal command completion
    event.completed();
  }
}

// Make function available to Office commands
window.rewriteToBusiness = rewriteToBusiness;


