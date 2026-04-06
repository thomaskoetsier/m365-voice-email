import { NextRequest, NextResponse } from "next/server";
import {
  getUnreadEmails,
  getEmailBody,
  sendReply,
  deleteEmail,
  markAsRead,
  decryptTokens,
} from "@/app/lib/outlook";

function getTokens(request: NextRequest) {
  const cookie = request.cookies.get("m365_tokens");
  if (!cookie) return null;
  try {
    return decryptTokens(cookie.value);
  } catch {
    return null;
  }
}

export async function POST(request: NextRequest) {
  const tokens = getTokens(request);
  if (!tokens) {
    return NextResponse.json(
      { error: "Not authenticated. Please connect Outlook first." },
      { status: 401 }
    );
  }

  const { action, ...params } = await request.json();

  try {
    switch (action) {
      case "list": {
        const emails = await getUnreadEmails(tokens, params.maxResults || 10);
        return NextResponse.json({ emails });
      }
      case "read": {
        const body = await getEmailBody(tokens, params.messageId);
        return NextResponse.json({ body });
      }
      case "reply": {
        await sendReply(tokens, params.messageId, params.body);
        return NextResponse.json({ success: true });
      }
      case "delete": {
        await deleteEmail(tokens, params.messageId);
        return NextResponse.json({ success: true });
      }
      case "markRead": {
        await markAsRead(tokens, params.messageId);
        return NextResponse.json({ success: true });
      }
      default:
        return NextResponse.json(
          { error: `Unknown action: ${action}` },
          { status: 400 }
        );
    }
  } catch (error: any) {
    console.error(`Outlook API error (${action}): ${error.message || "unknown"}`);
    return NextResponse.json(
      { error: "Outlook API error" },
      { status: 500 }
    );
  }
}
