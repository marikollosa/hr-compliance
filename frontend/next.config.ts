import type { NextConfig } from "next";

/** GitHub Pages project sites use /repo-name/. Set NEXT_PUBLIC_BASE_PATH in CI (see pages.yml). */
function normalizeBasePath(v: string | undefined): string {
  if (!v?.trim()) return "";
  const s = v.trim().replace(/\/+$/, "");
  return s.startsWith("/") ? s : `/${s}`;
}

const basePath = normalizeBasePath(
  process.env.NEXT_PUBLIC_BASE_PATH ?? process.env.BASE_PATH
);

const nextConfig: NextConfig = {
  // Required for GitHub Pages static hosting
  output: "export",

  basePath,
  assetPrefix: basePath ? `${basePath}/` : "",

  // GitHub Pages does not support Next Image optimization
  images: {
    unoptimized: true,
  },
};

export default nextConfig;
