import type { NextConfig } from "next";

const addInHeaders = [
  { key: "Access-Control-Allow-Origin", value: "*" },
  {
    key: "Access-Control-Allow-Headers",
    value: "Authorization, Content-Type",
  },
  { key: "Access-Control-Allow-Methods", value: "GET, POST, OPTIONS" },
  { key: "Content-Security-Policy", value: "frame-ancestors *" },
];

const nextConfig: NextConfig = {
  async rewrites() {
    return [{ source: "/manifest.xml", destination: "/manifest" }];
  },
  async headers() {
    return [
      { source: "/manifest", headers: addInHeaders },
      { source: "/manifest.xml", headers: addInHeaders },
      { source: "/taskpane", headers: addInHeaders },
      { source: "/taskpane/:path*", headers: addInHeaders },
      { source: "/api/captures", headers: addInHeaders },
      { source: "/api/logs", headers: addInHeaders },
      { source: "/api/letter-health", headers: addInHeaders },
      { source: "/api/letter-config", headers: addInHeaders },
      { source: "/api/letter-client-result", headers: addInHeaders },
      { source: "/api/deploy-info", headers: addInHeaders },
      { source: "/launchevent.js", headers: addInHeaders },
      { source: "/commands.html", headers: addInHeaders },
      { source: "/commands.js", headers: addInHeaders },
      { source: "/capture.js", headers: addInHeaders },
      { source: "/taskpane.html", headers: addInHeaders },
      { source: "/taskpane.js", headers: addInHeaders },
      { source: "/simulator.html", headers: addInHeaders },
      { source: "/icons/:path*", headers: addInHeaders },
    ];
  },
};

export default nextConfig;
