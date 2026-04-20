import type { RequestHandler, Response } from "express";
import { timingSafeEqual } from "node:crypto";
import { createRemoteJWKSet, jwtVerify } from "jose";

const TENANT_ID = process.env.AAD_TENANT_ID ?? "2131d362-e12d-44c1-843b-a1413d6b96a3";
const AUDIENCE = process.env.AAD_API_AUDIENCE ?? "api://7e277f47-2b3e-432d-aad1-01cfe820b5e2";
const API_KEY = process.env.MCP_API_KEY ?? "";

const JWKS = createRemoteJWKSet(
  new URL(`https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`),
  { cooldownDuration: 30_000, cacheMaxAge: 6 * 60 * 60 * 1000 },
);

// AAD emits v1 issuer (sts.windows.net) for client_credentials + .default scope,
// and v2 issuer when accessTokenAcceptedVersion=2 on the app manifest. Accept both.
const ALLOWED_ISSUERS = new Set([
  `https://sts.windows.net/${TENANT_ID}/`,
  `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
]);

const BEARER_PATTERN = /^Bearer\s+(.+)$/i;

function unauthorized(res: Response, msg: string) {
  res.setHeader(
    "WWW-Authenticate",
    `Bearer error="invalid_token", error_description="${msg.replace(/"/g, "'")}"`,
  );
  res.status(401).json({ error: "unauthorized", detail: msg });
}

function forbidden(res: Response, msg: string) {
  res.status(403).json({ error: "forbidden", detail: msg });
}

function constantTimeMatch(presented: string, expected: string): boolean {
  if (!expected || !presented) return false;
  const a = Buffer.from(presented);
  const b = Buffer.from(expected);
  if (a.length !== b.length) return false;
  return timingSafeEqual(a, b);
}

export function requireRole(requiredRole: string): RequestHandler {
  return async (req, res, next) => {
    // Path 1: static X-API-Key header (used by Power Platform custom connector)
    // Only accepted when MCP_API_KEY env is set.
    const presentedKey = req.get("x-api-key");
    if (presentedKey && API_KEY && constantTimeMatch(presentedKey, API_KEY)) {
      (req as unknown as { auth: unknown }).auth = { type: "api-key" };
      return next();
    }

    // Path 2: AAD bearer JWT with required role (used by Claude Desktop proxy etc.)
    const header = req.get("authorization") ?? "";
    const matched = header.match(BEARER_PATTERN);
    if (!matched) return unauthorized(res, "missing bearer token or X-API-Key");

    try {
      const { payload } = await jwtVerify(matched[1], JWKS, {
        audience: AUDIENCE,
        algorithms: ["RS256"],
      });

      if (!ALLOWED_ISSUERS.has(String(payload.iss))) {
        return unauthorized(res, `unexpected issuer: ${payload.iss}`);
      }
      if (payload.tid !== TENANT_ID) {
        return unauthorized(res, "wrong tenant");
      }

      const roles = Array.isArray(payload.roles) ? (payload.roles as string[]) : [];
      if (!roles.includes(requiredRole)) {
        return forbidden(res, `missing required role: ${requiredRole}`);
      }

      (req as unknown as { auth: unknown }).auth = {
        type: "jwt",
        sub: payload.sub,
        appid: payload.appid,
        roles,
      };
      next();
    } catch (err) {
      const detail =
        err && typeof err === "object" && "code" in err
          ? String((err as { code: unknown }).code)
          : err instanceof Error
            ? err.message
            : "unknown";
      return unauthorized(res, `token verification failed: ${detail}`);
    }
  };
}
