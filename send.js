import puppeteer from "puppeteer";
import xlsx from "xlsx";
import fs from "fs";
import { exec } from "child_process";
import path from "path";

const filePath = "./data/startup_phone.xlsx";
const countryCode = "91";
const imagePath = path.resolve("./");

if (!fs.existsSync(filePath) || !fs.existsSync(imagePath)) {
  console.error("⚠️ Missing required files! Ensure the Excel and image exist.");
  process.exit(1);
}

const readExcelData = (filePath) => {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return xlsx.utils.sheet_to_json(sheet);
};

const formatMessage = (fullName, companyName) => {
  return [
    `Hi *${fullName?.trim() || ""}*,`,
    ``,
    `Just thought to reshare the registration link in case you were not able to apply last time.`,
    ``,
    `Click on the link below or scan the QR Code to apply.`,
    ``,
    `https://lu.ma/yphrrm1k`,
    ``,
    `Please feel free to contact us in case of any queries`,
    ``,
    `Regards `,
    `Team PedalStart`,
  ];
};

const copyImageToClipboard = async () => {
  return new Promise((resolve, reject) => {
    const command =
      process.platform === "win32"
        ? `powershell -command "Set-Clipboard -Path '${imagePath}'"`
        : `xclip -selection clipboard -t image/png -i '${imagePath}'`;

    exec(command, (error) => {
      if (error) {
        console.error("❌ Error copying image to clipboard:", error.message);
        reject(error);
      } else {
        console.log("📸 Image copied to clipboard successfully.");
        resolve();
      }
    });
  });
};

const sendWhatsAppMessages = async (contacts) => {
  const browser = await puppeteer.launch({ headless: false });
  const page = await browser.newPage();
  await page.goto("https://web.whatsapp.com/", { timeout: 60000 });

  console.log("📲 Scan QR Code on WhatsApp Web and press ENTER.");
  await new Promise((resolve) => process.stdin.once("data", resolve));

  for (let contact of contacts) {
    const { Phone, fullName, companyName } = contact;
    if (!Phone) continue;

    const messageLines = formatMessage(fullName, companyName);

    try {
      await page.goto(
        `https://web.whatsapp.com/send?phone=${countryCode}${Phone}`,
        { timeout: 60000 }
      );
      await page.waitForSelector(
        'div[aria-label="Type a message"][role="textbox"]',
        { timeout: 20000 }
      );

      await page.evaluate(async (lines) => {
        const textBox = document.querySelector(
          'div[aria-label="Type a message"][role="textbox"]'
        );
        if (!textBox) return;

        textBox.focus();

        for (let i = 0; i < lines.length; i++) {
          const line = lines[i].trim();
          if (line) {
            document.execCommand("insertText", false, line);
          }
          // Always add a line break, even for empty lines
          const event = new KeyboardEvent("keydown", {
            key: "Enter",
            code: "Enter",
            keyCode: 13,
            which: 13,
            bubbles: true,
            cancelable: true,
            shiftKey: true, // Using Shift+Enter for line break without sending
          });
          textBox.dispatchEvent(event);
        }
      }, messageLines);

      console.log(`💬 Message formatted correctly for: ${Phone}`);

      await new Promise((res) => setTimeout(res, 3000));

      // Copy and paste image
      await copyImageToClipboard();
      await page.keyboard.down("Control");
      await page.keyboard.press("V");
      await page.keyboard.up("Control");
      console.log(`📤 Image pasted from clipboard for: ${Phone}`);

      await new Promise((res) => setTimeout(res, 5000));

      // Send message
      const sendButtonSelector = 'button span[data-icon="send"]';
      await page.waitForSelector(sendButtonSelector, { timeout: 10000 });
      await page.click(sendButtonSelector);
      console.log(`✅ Message sent with image to: ${Phone}`);
    } catch (error) {
      console.error(`❌ Error processing ${Phone}: ${error.message}`);
    }

    await new Promise((res) => setTimeout(res, 5000));
  }

  console.log("🎉 Process completed!");
  await browser.close();
};

const contacts = readExcelData(filePath);
sendWhatsAppMessages(contacts);
