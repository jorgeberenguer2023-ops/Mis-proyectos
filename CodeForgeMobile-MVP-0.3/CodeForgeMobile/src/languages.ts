export type LanguageDefinition = { id: string; label: string; extensions: string[] };
export const languages: LanguageDefinition[] = [
  { id:"javascript", label:"JavaScript", extensions:[".js",".jsx"] },
  { id:"typescript", label:"TypeScript", extensions:[".ts",".tsx"] },
  { id:"python", label:"Python", extensions:[".py"] },
  { id:"java", label:"Java", extensions:[".java"] },
  { id:"c", label:"C", extensions:[".c",".h"] },
  { id:"cpp", label:"C++", extensions:[".cpp",".hpp",".cc",".cxx"] },
  { id:"csharp", label:"C#", extensions:[".cs"] },
  { id:"php", label:"PHP", extensions:[".php"] },
  { id:"go", label:"Go", extensions:[".go"] },
  { id:"rust", label:"Rust", extensions:[".rs"] },
  { id:"kotlin", label:"Kotlin", extensions:[".kt",".kts"] },
  { id:"swift", label:"Swift", extensions:[".swift"] },
  { id:"dart", label:"Dart", extensions:[".dart"] },
  { id:"ruby", label:"Ruby", extensions:[".rb"] },
  { id:"sql", label:"SQL", extensions:[".sql"] },
  { id:"html", label:"HTML", extensions:[".html",".htm"] },
  { id:"css", label:"CSS", extensions:[".css"] },
  { id:"scss", label:"SCSS", extensions:[".scss"] },
  { id:"bash", label:"Shell / Bash", extensions:[".sh",".bash"] },
  { id:"powershell", label:"PowerShell", extensions:[".ps1"] },
  { id:"r", label:"R", extensions:[".r",".R"] },
  { id:"matlab", label:"MATLAB", extensions:[".m"] },
  { id:"lua", label:"Lua", extensions:[".lua"] },
  { id:"perl", label:"Perl", extensions:[".pl",".pm"] },
  { id:"scala", label:"Scala", extensions:[".scala"] },
  { id:"objective-c", label:"Objective-C", extensions:[".mm",".m"] },
  { id:"haskell", label:"Haskell", extensions:[".hs"] },
  { id:"elixir", label:"Elixir", extensions:[".ex",".exs"] },
  { id:"clojure", label:"Clojure", extensions:[".clj",".cljs"] },
  { id:"julia", label:"Julia", extensions:[".jl"] },
  { id:"yaml", label:"YAML", extensions:[".yaml",".yml"] },
  { id:"json", label:"JSON", extensions:[".json"] },
  { id:"xml", label:"XML", extensions:[".xml"] },
  { id:"markdown", label:"Markdown", extensions:[".md",".markdown"] },
  { id:"dockerfile", label:"Dockerfile", extensions:["Dockerfile"] }
];
export function detectLanguage(filename: string): LanguageDefinition | null {
  const name = filename.toLowerCase();
  return languages.find(l => l.extensions.some(ext =>
    ext.startsWith(".") ? name.endsWith(ext.toLowerCase()) : name === ext.toLowerCase()
  )) ?? null;
}