// Ebook Vault – অর্ডার রেকর্ড + এক-ক্লিকে অ্যাপ্রুভ + অটো ডেলিভারি ইমেইল
// Google Sheet এর Extensions > Apps Script এ পেস্ট করুন।

const OWNER_EMAIL   = 'minideveloper41@gmail.com';
const DELIVERY_LINK = 'https://drive.google.com/drive/folders/1KGnF6Ob_ngI5w-bTKjdjOUgrhis5VJoC'; // শুধু যে বইগুলো বিক্রির অধিকার আপনার আছে সেগুলোর ফোল্ডার লিংক
const STORE_NAME    = 'Ebook Vault';
const SUPPORT_WA    = '01941961084';

const HEAD = ['OrderID','Time','Name','Email','Phone','Method','Sender','TrxID','Status','Token'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Orders');
  if (!sh) { sh = ss.insertSheet('Orders'); sh.appendRow(HEAD); }
  return sh;
}

// ওয়েবসাইট থেকে অর্ডার আসে
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d.name || !d.email || !d.trxid || !/^\S+@\S+\.\S+$/.test(d.email)) return out_('bad');
    const sh = sheet_();
    const id = 'EV' + Utilities.formatDate(new Date(), 'Asia/Dhaka', 'yyMMddHHmmss');
    const token = Utilities.getUuid();
    const time = Utilities.formatDate(new Date(), 'Asia/Dhaka', 'dd MMM yyyy, hh:mm a');
    sh.appendRow([id, time, d.name, d.email, d.phone, d.method, d.sender, d.trxid, 'Pending', token]);

    const approve = ScriptApp.getService().getUrl() + '?action=approve&id=' + id + '&t=' + token;
    const html =
      '<div style="font-family:Arial,sans-serif;max-width:520px">' +
      '<h2>নতুন অর্ডার ' + id + '</h2>' +
      '<p><b>নাম:</b> ' + esc_(d.name) + '<br><b>ইমেইল:</b> ' + esc_(d.email) +
      '<br><b>ফোন:</b> ' + esc_(d.phone) + '<br><b>মাধ্যম:</b> ' + esc_(d.method) +
      '<br><b>যে নম্বর থেকে:</b> ' + esc_(d.sender) + '<br><b>TrxID:</b> ' + esc_(d.trxid) + '<br><b>পরিমাণ:</b> ৳59</p>' +
      '<p>bKash/Nagad অ্যাপে TrxID মিলিয়ে নিন, তারপর:</p>' +
      '<p><a href="' + approve + '" style="background:#ffc83d;color:#1a1200;padding:14px 26px;border-radius:10px;text-decoration:none;font-weight:bold;display:inline-block">✅ Continue – ডেলিভারি পাঠান</a></p>' +
      '</div>';
    MailApp.sendEmail({ to: OWNER_EMAIL, subject: 'নতুন অর্ডার – ' + d.name + ' (' + id + ')', htmlBody: html });
    return out_('ok');
  } catch (err) { return out_('error'); }
}

// আপনি "Continue" লিংকে ক্লিক করলে
function doGet(e) {
  const p = e.parameter || {};
  if (p.action !== 'approve') return page_('Ebook Vault order system চালু আছে।');
  const sh = sheet_();
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === p.id && rows[i][9] === p.t) {
      if (rows[i][8] === 'Sent') return page_('এই অর্ডারের বই আগেই পাঠানো হয়েছে।');
      const name = rows[i][2], email = rows[i][3];
      MailApp.sendEmail({
        to: email,
        subject: '🎉 আপনার ই-বুক কম্বো প্যাক – ' + STORE_NAME,
        htmlBody: customerMail_(name, p.id),
        name: STORE_NAME
      });
      sh.getRange(i + 1, 9).setValue('Sent');
      return page_('✅ ' + name + ' কে (' + email + ') ডেলিভারি ইমেইল পাঠানো হয়েছে।');
    }
  }
  return page_('অর্ডার পাওয়া যায়নি বা লিংক ভুল।');
}

function customerMail_(name, id) {
  return '<div style="background:#06140d;padding:24px;font-family:Arial,sans-serif">' +
    '<div style="max-width:540px;margin:auto;background:#0d2217;border:1px solid #1f4a33;border-radius:18px;padding:30px;color:#eef7f0">' +
    '<div style="font-size:24px;font-weight:bold">Ebook <span style="color:#2fe06f">Vault</span></div>' +
    '<h2 style="color:#ffc83d;margin:18px 0 8px">ধন্যবাদ, ' + esc_(name) + '!</h2>' +
    '<p style="line-height:1.7;color:#cfe3d5">আপনার পেমেন্ট যাচাই হয়েছে। নিচের বাটনে ক্লিক করে আপনার ই-বুক কম্বো প্যাক ডাউনলোড করুন।</p>' +
    '<p style="margin:26px 0"><a href="' + DELIVERY_LINK + '" style="background:#ffc83d;color:#1a1200;padding:15px 30px;border-radius:12px;text-decoration:none;font-weight:bold;font-size:17px;display:inline-block">📥 কম্বো প্যাক ডাউনলোড করুন</a></p>' +
    '<p style="color:#9bb8a6;font-size:14px;line-height:1.7">অর্ডার আইডি: ' + id + '<br>ডাউনলোডে সমস্যা হলে হোয়াটসঅ্যাপ করুন: ' + SUPPORT_WA + '</p>' +
    '</div></div>';
}

function esc_(s) { return String(s || '').replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function out_(s) { return ContentService.createTextOutput(s); }
function page_(m) {
  return HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:40px;text-align:center;font-size:20px">' + m + '</div>');
}
