export type GitHubUser = { login: string; name?: string | null; avatar_url?: string };
export type GitHubRepository = {
  id: number; name: string; full_name: string; private: boolean;
  default_branch: string; html_url: string; owner: { login: string };
};
export type GitHubBranch = { name: string; commit: { sha: string } };
export type GitHubContent = {
  name: string; path: string; sha: string; type: string;
  content?: string; encoding?: string; download_url?: string | null;
};
export type GitHubCommitResult = { content: GitHubContent | null; commit: { sha: string; html_url: string } };
export type GitHubCompare = {
  status: string; ahead_by: number; behind_by: number;
  total_commits: number;
  files?: Array<{ filename: string; status: string; additions: number; deletions: number; patch?: string }>;
};