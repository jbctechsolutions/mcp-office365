# syntax=docker/dockerfile:1
#
# Container image for remote connector mode (`serve`). Multi-stage so the native
# better-sqlite3 addon is compiled against the runtime's exact Node ABI and baked
# in — the image never hits the NODE_MODULE_VERSION mismatch a stale prebuild
# would. Debian slim (glibc) is used over Alpine (musl) for reliable native builds.

# ---- builder: install deps, compile the native addon, build TS ----
FROM node:26-bookworm-slim AS builder
WORKDIR /app

# Toolchain for better-sqlite3's node-gyp build. Removed with the stage.
RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 make g++ \
  && rm -rf /var/lib/apt/lists/*

# Install with the lockfile for reproducibility (compiles better-sqlite3).
COPY package.json package-lock.json ./
RUN npm ci

# Build the TypeScript (tsc + build-info + chmod).
COPY tsconfig.json ./
COPY scripts ./scripts
COPY src ./src
RUN npm run build

# Drop dev dependencies, keeping the compiled better-sqlite3 binary.
RUN npm prune --omit=dev

# ---- runtime: minimal, non-root ----
FROM node:26-bookworm-slim AS runtime
ENV NODE_ENV=production
WORKDIR /app

# Unprivileged user (node images ship a `node` user uid 1000).
COPY --from=builder --chown=node:node /app/node_modules ./node_modules
COPY --from=builder --chown=node:node /app/dist ./dist
COPY --chown=node:node package.json ./

USER node
EXPOSE 8080

# Liveness via the server's own /healthz (no curl in the image — use node fetch).
HEALTHCHECK --interval=30s --timeout=3s --start-period=10s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:8080/healthz').then(r=>process.exit(r.ok?0:1)).catch(()=>process.exit(1))"

# Remote connector mode. Host/port here; all auth/state config comes from env
# (see docs/remote/deployment.md §2). Bind 0.0.0.0 — auth is required whenever
# OUTLOOK_MCP_CONNECTOR_URL is set, and the server refuses a non-loopback bind
# without it.
ENTRYPOINT ["node", "dist/index.js"]
CMD ["serve", "--host", "0.0.0.0", "--port", "8080"]
