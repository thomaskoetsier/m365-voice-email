import { NextRequest, NextResponse } from "next/server";
import { getAuthUrl } from "@/app/lib/outlook";

export async function GET(request: NextRequest) {
  const cookie = request.cookies.get("m365_tokens");
  if (cookie) {
    return NextResponse.json({ authenticated: true });
  }
  const url = getAuthUrl();
  return NextResponse.redirect(url);
}
