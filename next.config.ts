import type { NextConfig } from "next";

const addInHeaders = [
  { key: "Access-Control-Allow-Origin", value: "*" },
  {
    key: "Access-Control-Allow-Headers",
    value: "Authorization, Content-Type",
  },
  { key: "Access-Control-Allow-Methods", value: "GET, OPTIONS" },
  { key: "Content-Security-Policy", value: "frame-ancestors *" },
];

const nextConfig: NextConfig = {
  async headers() {
    return [
      { source: "/taskpane", headers: addInHeaders },
      { source: "/taskpane/:path*", headers: addInHeaders },
      { source: "/manifest.xml", headers: addInHeaders },
      { source: "/launchevent.js", headers: addInHeaders },
      { source: "/commands.html", headers: addInHeaders },
      { source: "/taskpane.html", headers: addInHeaders },
      { source: "/taskpane.js", headers: addInHeaders },
      { source: "/simulator.html", headers: addInHeaders },
      { source: "/icons/:path*", headers: addInHeaders },
    ];
  },
};

export default nextConfig;
