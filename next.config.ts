import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  reactCompiler: true,
  allowedDevOrigins: ["127.0.0.1", "localhost"],
  async headers() {
    return [
      {
        source: "/(.*)",
        headers: [
          {
            key: "Content-Security-Policy",
            value:
              "default-src 'self' https://appsforoffice.microsoft.com; script-src 'self' 'unsafe-inline' 'unsafe-eval' https://appsforoffice.microsoft.com; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; connect-src 'self' ws: wss: https:;",
          },
        ],
      },
    ];
  },
};

export default nextConfig;
