import * as SecureStore from "expo-secure-store";
import type { GitHubBranch, GitHubCompare, GitHubContent, GitHubRepository, GitHubUser, GitHubCommitResult } from "./types";

const TOKEN_KEY = "codeforge.github.token";
const API = "https://api.github.com";

export async function saveGitHubToken(token: string) {
  const clean = token.trim();
  if (!clean) throw new Error("Introduce un token de GitHub.");
  await SecureStore.setItemAsync(TOKEN_KEY, clean);
}
export async function getGitHubToken() {
  return SecureStore.getItemAsync(TOKEN_KEY);
}
export async function clearGitHubToken() {
  await SecureStore.deleteItemAsync(TOKEN_KEY);
}
async function request<T>(path: string, init: RequestInit = {}): Promise<T> {
  const token = await getGitHubToken();
  if (!token) throw new Error("Primero guarda un token de GitHub en la pestaña GitHub.");
  const response = await fetch(`${API}${path}`, {
    ...init,
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${token}`,
      "X-GitHub-Api-Version": "2022-11-28",
      ...(init.body ? { "Content-Type": "application/json" } : {}),
      ...init.headers
    }
  });
  if (!response.ok) {
    const message = await response.text().catch(() => "");
    throw new Error(`GitHub API ${response.status}: ${message.slice(0, 240) || response.statusText}`);
  }
  if (response.status === 204) return undefined as T;
  return response.json() as Promise<T>;
}
export const github = {
  user: () => request<GitHubUser>("/user"),
  repositories: () => request<GitHubRepository[]>("/user/repos?sort=updated&per_page=100&affiliation=owner,collaborator,organization_member"),
  branches: (owner: string, repo: string) => request<GitHubBranch[]>(`/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/branches?per_page=100`),
  contents: (owner: string, repo: string, path: string, ref?: string) =>
    request<GitHubContent>(`/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/${path.split("/").map(encodeURIComponent).join("/")}${ref ? `?ref=${encodeURIComponent(ref)}` : ""}`),
  commitFile: (owner: string, repo: string, path: string, message: string, content: string, branch: string, sha?: string) =>
    request<GitHubCommitResult>(`/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/${path.split("/").map(encodeURIComponent).join("/")}`, {
      method: "PUT",
      body: JSON.stringify({ message, content: toBase64Utf8(content), branch, ...(sha ? { sha } : {}) })
    }),
  compare: (owner: string, repo: string, base: string, head: string) =>
    request<GitHubCompare>(`/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/compare/${encodeURIComponent(base)}...${encodeURIComponent(head)}`)
};
function toBase64Utf8(value: string): string {
  // Encode UTF-8 without relying on browser btoa, which may not exist in React Native.
  const bytes = new TextEncoder().encode(value);
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  let out = "";
  for (let i = 0; i < bytes.length; i += 3) {
    const a = bytes[i], hasB = i + 1 < bytes.length, hasC = i + 2 < bytes.length;
    const b = hasB ? bytes[i + 1] : 0, c = hasC ? bytes[i + 2] : 0;
    out += alphabet[a >> 2];
    out += alphabet[((a & 3) << 4) | (b >> 4)];
    out += hasB ? alphabet[((b & 15) << 2) | (c >> 6)] : "=";
    out += hasC ? alphabet[c & 63] : "=";
  }
  return out;
}