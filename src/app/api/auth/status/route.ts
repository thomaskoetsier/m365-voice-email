import { NextRequest, NextResponse } from "next/server";

export async function GET(request: NextRequest) {
  const cookie = request.cookies.get("m365_tokens");
  return NextResponse.json({ authenticated: !!cookie });
}
