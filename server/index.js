import express from "express";
import cors from "cors";
import path from "path";
import { fileURLToPath } from "url";
import OpenAI from "openai";
import { promises as fs } from "fs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json({ limit: "1mb" }));

// Serve static web assets
app.use("/web", express.static(path.join(__dirname, "../web")));

// Tiny transparent PNG as placeholder icons
const transparentPngBase64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAoMBgUuK2y8AAAAASUVORK5CYII=";
app.get(["/web/assets/icon-16.png", "/web/assets/icon-32.png", "/web/assets/icon-80.png"], (_req, res) => {
  res.type("png").send(Buffer.from(transparentPngBase64, "base64"));
});

const port = process.env.PORT ? Number(process.env.PORT) : 3000;
const publicBaseUrl = process.env.PUBLIC_BASE_URL; // optional explicit base URL

// OpenAI client
const openai = new OpenAI({ apiKey: process.env.OPENAI_API_KEY });
const preferredModel = process.env.OPENAI_MODEL || "gpt-5"; // per requirement
const fallbackModel = process.env.OPENAI_FALLBACK_MODEL || "gpt-4o-mini";

// Prompt loading (externalized to 10_Assets/propmt)
const promptFileEnv = process.env.PROMPT_FILE;
const promptFileDefault = path.join(__dirname, "../AI_setting/system_prompt.md");
const defaultSystemPrompt =
  "あなたは日本語のビジネスメール作成の専門家です。ユーザーが書いた文章を、" +
  "丁寧で簡潔、読みやすいビジネス文書へ言い換えてください。意味は変えず、敬語・語調を整え、" +
  "必要に応じて件名候補を1行目に [件名] として付与してください。出力は本文のみ。";

async function readSystemPrompt() {
  const filePath = promptFileEnv || promptFileDefault;
  try {
    const text = await fs.readFile(filePath, "utf8");
    return (text || "").trim() || defaultSystemPrompt;
  } catch {
    return defaultSystemPrompt;
  }
}

app.get("/", (_req, res) => {
  res.type("text/plain").send("OK");
});

app.post("/api/rewrite", async (req, res) => {
  try {
    const { text } = req.body || {};
    if (!text || typeof text !== "string") {
      return res.status(400).json({ error: "text is required" });
    }

    const systemPrompt = await readSystemPrompt();

    async function createCompletion(model, prompt) {
      const completion = await openai.chat.completions.create({
        model,
        temperature: 0.3,
        messages: [
          { role: "system", content: prompt },
          { role: "user", content: text }
        ],
        max_tokens: 1200
      });
      const result = completion.choices?.[0]?.message?.content?.trim();
      if (!result) throw new Error("No content returned from model");
      return result;
    }

    let result;
    try {
      result = await createCompletion(preferredModel, systemPrompt);
    } catch (e) {
      // Fallback to a widely available model if preferred fails
      result = await createCompletion(fallbackModel, systemPrompt);
    }

    res.json({ result });
  } catch (err) {
    res.status(500).json({ error: err?.message || "internal error" });
  }
});

// Dynamic manifest.xml to reflect current base URL
app.get("/manifest.xml", (req, res) => {
  const baseUrl = publicBaseUrl || `${req.protocol}://${req.get("host")}`;
  res.type("application/xml").send(generateManifestXml(baseUrl));
});

function generateManifestXml(baseUrl) {
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="MailApp">
  <Id>b2a8e5f8-5c74-4b9a-9b1f-1dbe9f07d111</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Pragmateches</ProviderName>
  <DefaultLocale>ja-JP</DefaultLocale>
  <DisplayName DefaultValue="ビジネス文面リライト" />
  <Description DefaultValue="下書きをビジネス向けにリライトします" />
  <IconUrl DefaultValue="${baseUrl}/web/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="${baseUrl}/web/assets/icon-80.png" />
  <SupportUrl DefaultValue="${baseUrl}" />
  <AppDomains>
    <AppDomain>${baseUrl}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Mailbox" />
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="Mailbox" />
    </Sets>
  </Requirements>
  <FormSettings>
    <Form xsi:type="ItemEdit">
      <DesktopSettings>
        <SourceLocation DefaultValue="${baseUrl}/web/function-file.html" />
      </DesktopSettings>
    </Form>
  </FormSettings>
  <Permissions>ReadWriteItem</Permissions>

  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" />
  </Rule>

  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" Version="1.0">
    <Hosts>
      <Host xsi:type="MailHost">
        <Runtimes>
          <Runtime resid="functionfile" lifetime="long" />
        </Runtimes>
        <ExtensionPoint xsi:type="MessageComposeCommandSurface">
          <OfficeTab id="TabDefault">
            <Group id="grpCompose" label="リライト">
              <Control xsi:type="Button" id="btnRewrite">
                <Label resid="btnRewriteLabel" />
                <Supertip>
                  <Title resid="btnRewriteLabel" />
                  <Description resid="btnRewriteDesc" />
                </Supertip>
                <Icon>
                  <bt:Image size="16" resid="icon16" />
                  <bt:Image size="32" resid="icon32" />
                  <bt:Image size="80" resid="icon80" />
                </Icon>
                <Action xsi:type="ExecuteFunction">
                  <FunctionName>rewriteToBusiness</FunctionName>
                </Action>
              </Control>
            </Group>
          </OfficeTab>
        </ExtensionPoint>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon16" DefaultValue="${baseUrl}/web/assets/icon-16.png" />
        <bt:Image id="icon32" DefaultValue="${baseUrl}/web/assets/icon-32.png" />
        <bt:Image id="icon80" DefaultValue="${baseUrl}/web/assets/icon-80.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="functionfile" DefaultValue="${baseUrl}/web/function-file.html" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="btnRewriteLabel" DefaultValue="ビジネスに整える" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="btnRewriteDesc" DefaultValue="下書きを丁寧で読みやすいビジネス文にリライトします" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>`;
}

app.listen(port, () => {
  console.log(`Server listening on http://0.0.0.0:${port}`);
});


