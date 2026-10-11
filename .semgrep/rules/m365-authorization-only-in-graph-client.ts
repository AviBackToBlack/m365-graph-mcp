declare const token: string;
declare const headers: Headers;
declare const record: Record<string, string>;
declare const name: string;

function build(): void {
  // ruleid: m365-authorization-only-in-graph-client
  const a = { Authorization: `Bearer ${token}` };
  // ruleid: m365-authorization-only-in-graph-client
  const b = { authorization: token };
  // ruleid: m365-authorization-only-in-graph-client
  const c = { "Authorization": token };
  // ruleid: m365-authorization-only-in-graph-client
  const c2 = { AUTHORIZATION: token };
  // ruleid: m365-authorization-only-in-graph-client
  headers.set("Authorization", token);
  // ruleid: m365-authorization-only-in-graph-client
  headers.append("authorization", token);
  // ruleid: m365-authorization-only-in-graph-client
  headers.set("AUTHORIZATION", token);
  // ruleid: m365-authorization-only-in-graph-client
  record.Authorization = token;
  // ruleid: m365-authorization-only-in-graph-client
  record.authorization = token;
  // ruleid: m365-authorization-only-in-graph-client
  record["Authorization"] = token;
  // ruleid: m365-authorization-only-in-graph-client
  record["AUTHORIZATION"] = token;
  // ruleid: m365-authorization-only-in-graph-client
  const d = `Bearer ${token}`;
  // ruleid: m365-authorization-only-in-graph-client
  const e = "Bearer " + token;

  // ok: m365-authorization-only-in-graph-client
  const f = { Accept: "application/json" };
  // ok: m365-authorization-only-in-graph-client
  headers.set("ConsistencyLevel", "eventual");
  // ok: m365-authorization-only-in-graph-client
  record["Prefer"] = "outlook.timezone";
  // ok: m365-authorization-only-in-graph-client
  const g = name.toLowerCase() === "authorization";
  // ok: m365-authorization-only-in-graph-client
  const r = record["Authorization"];
  // ok: m365-authorization-only-in-graph-client
  const h = "Bearer [redacted]";
  // ok: m365-authorization-only-in-graph-client
  const i = /^Bearer ([^\s]+)$/i.exec(token);
  void [a, b, c, c2, d, e, f, g, r, h, i];
}
build();
