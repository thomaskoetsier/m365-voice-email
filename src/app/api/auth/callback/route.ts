import { NextRequest, NextResponse } from "next/server";
import { exchangeCode, encryptTokens } from "@/app/lib/outlook";

export async function GET(request: NextRequest) {
  const code = request.nextUrl.searchParams.get("code");
  if (!code) {
    return NextResponse.json({ error: "No code provided" }, { status: 400 });
  }

  try {
    const tokenData = await exchangeCode(code);
    const tokens = {
      access_token: tokenData.access_token,
      refresh_token: tokenData.refresh_token,
      expires_at: Math.floor(Date.now() / 1000) + (tokenData.expires_in || 3600),
    };
    const encrypted = encryptTokens(tokens);

    const host = request.headers.get("x-forwarded-host") || request.headers.get("host") || "";
    const proto = request.headers.get("x-forwarded-proto") || "https";
    const origin = host ? `${proto}://${host}` : request.url;
    const response = NextResponse.redirect(new URL("/", origin));
    response.cookies.set("m365_tokens", encrypted, {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      maxAge: 60 * 60 * 24 * 30, // 30 days
      path: "/",
    });
    return response;
  } catch {
    console.error("OAuth callback error");
    return NextResponse.json(
      { error: "Authentication failed" },
      { status: 500 }
    );
  }
}
