import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  typescript:
    process.env.NEXT_BUILD_TARGET === "vercel"
      ? { tsconfigPath: "tsconfig.vercel.json" }
      : undefined,
};

export default nextConfig;
