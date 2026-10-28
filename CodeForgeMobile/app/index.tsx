import React, { useEffect, useMemo, useRef, useState } from "react";
import { SafeAreaView, View, Text, TextInput, Pressable, ScrollView, Alert, ActivityIndicator } from "react-native";
import AsyncStorage from "@react-native-async-storage/async-storage";
import { WebView } from "react-native-webview";
import { StatusBar } from "expo-status-bar";
import { languages, detectLanguage } from "../src/languages";
import { clearGitHubToken, getGitHubToken, github, saveGitHubToken } from "../src/github/client";
import type { GitHubBranch, GitHubRepository, GitHubUser } from "../src/github/types";

type Doc = { id: string; name: string; language: string; content: string; updatedAt: number };
const KEY = "codeforge.documents.v2";
const starter: Doc[] = [
  { id: "welcome", name: "hello.py", language: "python", content: 'def greet(name):\n    return f"Hola, {name}!"\n\nprint(greet("mundo"))\n', updatedAt: Date.now() },
  { id: "web", name: "index.html", language: "html", content: '<!doctype html>\n<html>\n  <body>\n    <h1>Hola, CodeForge</h1>\n  </body>\n</html>\n', updatedAt: Date.now() }
];
const C = { bg:"#0b1020", panel:"#111a2e", panel2:"#17233b", border:"#263653", text:"#edf3ff", muted:"#93a5c5", accent:"#7c9cff", green:"#4adea5", orange:"#ffbd69", red:"#ff7a90" };

function editorHtml(value: string, language: string) {
  const safe = JSON.stringify(value).replace(/</g, "\\u003c");
  const lang = JSON.stringify(language);
  return `<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1,maximum-scale=1"><style>
  *{box-sizing:border-box}body{margin:0;background:#0b1020;color:#e5edff;font:14px ui-monospace,SFMono-Regular,Consolas,monospace}
  .top{padding:9px 12px;color:#9aadd0;background:#111a2e;border-bottom:1px solid #263653;font:12px system-ui}
  textarea{display:block;width:100%;height:calc(100vh - 42px);resize:none;border:0;outline:0;background:transparent;color:#e5edff;padding:14px;line-height:1.65;tab-size:2;white-space:pre;overflow:auto;font:14px/1.65 ui-monospace,SFMono-Regular,Consolas,monospace}
  </style></head><body><div class="top">CODEFORGE EDITOR · <span id="lang"></span></div><textarea id="code" spellcheck="false" autocapitalize="off" autocomplete="off"></textarea>
  <script>
  const initial=${safe}; const language=${lang}; const el=document.getElementById('code'); el.value=initial; document.getElementById('lang').textContent=language.toUpperCase();
  function send(){ if(window.ReactNativeWebView) window.ReactNativeWebView.postMessage(JSON.stringify({type:'change',content:el.value})); }
  el.addEventListener('input',send);
  window.__runCode = function(){
    const output=[];
    const originalLog=console.log, originalWarn=console.warn, originalError=console.error;
    const fmt=(args)=>Array.from(args).map(v=>{try{return typeof v==='string'?v:JSON.stringify(v)}catch{return String(v)}}).join(' ');
    console.log=(...a)=>output.push(fmt(a)); console.warn=(...a)=>output.push('WARN: '+fmt(a)); console.error=(...a)=>output.push('ERROR: '+fmt(a));
    try { (new Function(el.value))(); }
    catch(err){ output.push((err&&err.name?err.name+': ':'Error: ')+(err&&err.message?err.message:String(err))); }
    finally { console.log=originalLog; console.warn=originalWarn; console.error=originalError; }
    if(window.ReactNativeWebView) window.ReactNativeWebView.postMessage(JSON.stringify({type:'runResult',output:output.join('\\n')||'(El código terminó sin salida de consola.)'}));
  };
  document.addEventListener('keydown',e=>{if(e.key==='Tab'){e.preventDefault();el.setRangeText('  ',el.selectionStart,el.selectionEnd,'end');send();}});
  </script></body></html>`;
}

export default function Home() {
  const [docs, setDocs] = useState<Doc[]>(starter);
  const [activeId, setActiveId] = useState("welcome");
  const [screen, setScreen] = useState<"editor"|"projects"|"github"|"languages"|"runner">("editor");
  const [search, setSearch] = useState("");
  const [draft, setDraft] = useState(starter[0].content);
  const [dirty, setDirty] = useState(false);
  const [busy, setBusy] = useState(false);
  const [notice, setNotice] = useState("");
  const [tokenInput, setTokenInput] = useState("");
  const [user, setUser] = useState<GitHubUser|null>(null);
  const [repos, setRepos] = useState<GitHubRepository[]>([]);
  const [selectedRepo, setSelectedRepo] = useState<GitHubRepository|null>(null);
  const [branches, setBranches] = useState<GitHubBranch[]>([]);
  const [branch, setBranch] = useState("");
  const [commitMessage, setCommitMessage] = useState("Update from CodeForge Mobile");
  const [remotePath, setRemotePath] = useState("");
  const [baseBranch, setBaseBranch] = useState("");
  const [compareBranch, setCompareBranch] = useState("");
  const [compareText, setCompareText] = useState("");
  const [runOutput, setRunOutput] = useState("");
  const editorRef = useRef<WebView>(null);
  const active = docs.find(d => d.id === activeId) ?? docs[0];

  useEffect(() => {
    AsyncStorage.getItem(KEY).then(raw => {
      if (raw) {
        const loaded = JSON.parse(raw) as Doc[];
        if (loaded.length) { setDocs(loaded); setActiveId(loaded[0].id); setDraft(loaded[0].content); }
      }
    }).catch(() => {});
    getGitHubToken().then(t => { if (t) setNotice("Token guardado de forma segura en este dispositivo."); }).catch(() => {});
  }, []);
  useEffect(() => { if (active) { setDraft(active.content); setDirty(false); } }, [activeId]);

  function runCurrentCode() {
    const lang = active?.language?.toLowerCase();
    if (!['javascript', 'js'].includes(lang)) {
      setRunOutput(`Ejecución local disponible solo para JavaScript por ahora. El archivo actual es ${lang || 'desconocido'}. Python y otros lenguajes necesitan un motor adicional que aún no está configurado.`);
      setScreen('runner');
      return;
    }
    setRunOutput('Ejecutando JavaScript…');
    editorRef.current?.injectJavaScript('window.__runCode && window.__runCode(); true;');
  }

  async function save(content = draft, quiet = false) {
    const next = docs.map(d => d.id === activeId ? { ...d, content, updatedAt: Date.now() } : d);
    setDocs(next); setDraft(content); setDirty(false);
    try {
      await AsyncStorage.setItem(KEY, JSON.stringify(next));
      if (!quiet) Alert.alert("Guardado", `${active.name} se ha guardado en este dispositivo.`);
    } catch { Alert.alert("Error", "No se pudo guardar el archivo."); }
  }
  function createFile() {
    const name = `nuevo-${docs.length + 1}.js`;
    const lang = detectLanguage(name)?.id ?? "javascript";
    const doc = { id: `${Date.now()}`, name, language: lang, content: "// Empieza a escribir aquí\n", updatedAt: Date.now() };
    const next = [...docs, doc]; setDocs(next); setActiveId(doc.id); setDraft(doc.content); setDirty(true); setScreen("editor");
    AsyncStorage.setItem(KEY, JSON.stringify(next)).catch(() => {});
  }
  async function withBusy(action: () => Promise<void>) {
    setBusy(true); setNotice("");
    try { await action(); } catch (e) { setNotice(e instanceof Error ? e.message : "Error inesperado."); }
    finally { setBusy(false); }
  }
  async function testGitHub() {
    await withBusy(async () => {
      const me = await github.user();
      setUser(me);
      const list = await github.repositories();
      setRepos(list);
      setNotice(`Conectado a GitHub como @${me.login}. Se cargaron ${list.length} repositorios.`);
    });
  }
  async function chooseRepo(repo: GitHubRepository) {
    setSelectedRepo(repo); setBranches([]); setBranch(""); setCompareText("");
    await withBusy(async () => {
      const list = await github.branches(repo.owner.login, repo.name);
      setBranches(list);
      setBranch(repo.default_branch);
      setBaseBranch(repo.default_branch);
      setCompareBranch(list.find(b => b.name !== repo.default_branch)?.name ?? repo.default_branch);
      setNotice(`Repositorio ${repo.full_name}: ${list.length} ramas cargadas.`);
    });
  }
  async function loadRemoteFile() {
    if (!selectedRepo || !branch || !remotePath.trim()) throw new Error("Selecciona repositorio, rama y ruta del archivo.");
    const file = await github.contents(selectedRepo.owner.login, selectedRepo.name, remotePath.trim(), branch);
    if (file.type !== "file" || !file.content) throw new Error("La ruta no es un archivo o GitHub no devolvió su contenido.");
    const content = decodeBase64(file.content.replace(/\n/g, ""));
    const name = file.name;
    const doc = { id: `gh-${Date.now()}`, name, language: detectLanguage(name)?.id ?? "plaintext", content, updatedAt: Date.now() };
    const next = [...docs, doc]; setDocs(next); setActiveId(doc.id); setDraft(content); setDirty(false); setScreen("editor");
    await AsyncStorage.setItem(KEY, JSON.stringify(next));
    setNotice(`Archivo ${file.path} importado desde ${selectedRepo.full_name}.`);
  }
  async function pushCurrentFile() {
    if (!selectedRepo || !branch || !active) throw new Error("Selecciona un repositorio y una rama.");
    if (!remotePath.trim()) throw new Error("Indica la ruta remota, por ejemplo src/main.py.");
    const old = await github.contents(selectedRepo.owner.login, selectedRepo.name, remotePath.trim(), branch).catch(() => null);
    const result = await github.commitFile(selectedRepo.owner.login, selectedRepo.name, remotePath.trim(), commitMessage.trim() || `Update ${remotePath}`, draft, branch, old?.sha);
    await save(draft, true);
    setNotice(`Commit creado: ${result.commit.sha.slice(0, 7)}. ${result.commit.html_url}`);
  }
  async function compareBranches() {
    if (!selectedRepo || !baseBranch || !compareBranch) throw new Error("Selecciona dos ramas.");
    const result = await github.compare(selectedRepo.owner.login, selectedRepo.name, baseBranch, compareBranch);
    const lines = [
      `${selectedRepo.full_name}: ${baseBranch} ← ${compareBranch}`,
      `Estado: ${result.status}; commits por delante: ${result.ahead_by}; por detrás: ${result.behind_by}`,
      `Commits totales: ${result.total_commits}`, "",
      ...(result.files ?? []).map(f => `${f.status.toUpperCase()} ${f.filename}  +${f.additions} / -${f.deletions}\n${f.patch ?? "(parche no disponible)"}`)
    ];
    setCompareText(lines.join("\n"));
    setNotice("Comparación obtenida de GitHub.");
  }
  async function removeToken() {
    await clearGitHubToken(); setUser(null); setRepos([]); setSelectedRepo(null); setBranches([]); setTokenInput("");
    setNotice("Token eliminado del almacenamiento seguro del dispositivo.");
  }

  const filtered = useMemo(() => languages.filter(l => l.label.toLowerCase().includes(search.toLowerCase())), [search]);

  return <SafeAreaView style={{ flex:1, backgroundColor:C.bg }}>
    <StatusBar style="light" />
    <View style={{ paddingHorizontal:16, paddingTop:10, paddingBottom:12, borderBottomWidth:1, borderBottomColor:C.border, flexDirection:"row", alignItems:"center", justifyContent:"space-between" }}>
      <View><Text style={{ color:C.accent, fontSize:11, fontWeight:"800", letterSpacing:2 }}>CODEFORGE</Text><Text style={{ color:C.text, fontSize:20, fontWeight:"800" }}>Mobile IDE <Text style={{ color:C.muted, fontSize:11 }}>MVP 0.3</Text></Text></View>
      <Pressable onPress={createFile} style={{ backgroundColor:C.accent, paddingHorizontal:13, paddingVertical:10, borderRadius:12 }}><Text style={{ color:"#081126", fontWeight:"800" }}>＋ Archivo</Text></Pressable>
    </View>
    <View style={{ flexDirection:"row", padding:7, gap:4, backgroundColor:C.panel }}>
      {([{id:"editor",label:"Editor"},{id:"projects",label:"Archivos"},{id:"languages",label:"Lenguajes"},{id:"github",label:"GitHub"},{id:"runner",label:"Ejecutar"}] as const).map(t =>
        <Pressable key={t.id} onPress={() => setScreen(t.id)} style={{ flex:1, paddingVertical:10, borderRadius:8, alignItems:"center", backgroundColor:screen===t.id?C.panel2:"transparent", borderWidth:screen===t.id?1:0, borderColor:C.border }}>
          <Text style={{ color:screen===t.id?C.text:C.muted, fontSize:10, fontWeight:"700" }}>{t.label}</Text>
        </Pressable>
      )}
    </View>
    {!!notice && <Pressable onPress={() => setNotice("")} style={{ marginHorizontal:10, marginTop:8, padding:9, backgroundColor:C.panel2, borderRadius:8 }}><Text style={{ color:C.muted, fontSize:11 }}>{notice}  ×</Text></Pressable>}
    {busy && <View style={{ flexDirection:"row", alignItems:"center", gap:8, paddingHorizontal:14, paddingTop:8 }}><ActivityIndicator color={C.accent}/><Text style={{ color:C.muted, fontSize:12 }}>Contactando servicio…</Text></View>}

    {screen==="editor" && active && <View style={{ flex:1 }}>
      <ScrollView horizontal showsHorizontalScrollIndicator={false} style={{ maxHeight:44, flexGrow:0, borderBottomWidth:1, borderBottomColor:C.border }}>
        {docs.map(d => <Pressable key={d.id} onPress={() => setActiveId(d.id)} style={{ paddingHorizontal:13, paddingVertical:12, backgroundColor:d.id===activeId?C.panel2:C.bg, borderRightWidth:1, borderRightColor:C.border }}>
          <Text style={{ color:d.id===activeId?C.text:C.muted, fontSize:12 }}>{d.name}{d.id===activeId&&dirty?" •":""}</Text>
        </Pressable>)}
      </ScrollView>
      <View style={{ flexDirection:"row", alignItems:"center", justifyContent:"space-between", padding:10 }}>
        <Pressable onPress={() => setScreen("languages")} style={{ backgroundColor:C.panel2, borderRadius:8, padding:8 }}><Text style={{ color:C.text, fontSize:12 }}>◈ {languages.find(l=>l.id===active.language)?.label??active.language} ▾</Text></Pressable>
        <View style={{ flexDirection:"row", gap:7 }}><Pressable onPress={runCurrentCode} style={{ backgroundColor:C.accent, paddingHorizontal:12, paddingVertical:9, borderRadius:8 }}><Text style={{ color:"#081126", fontWeight:"800" }}>▶ Ejecutar</Text></Pressable><Pressable onPress={() => save()} style={{ backgroundColor:C.green, paddingHorizontal:12, paddingVertical:9, borderRadius:8 }}><Text style={{ color:"#062116", fontWeight:"800" }}>Guardar</Text></Pressable></View>
      </View>
      <View style={{ flex:1, marginHorizontal:8, marginBottom:8, borderWidth:1, borderColor:C.border, borderRadius:12, overflow:"hidden" }}>
        <WebView ref={editorRef} originWhitelist={["*"]} source={{ html:editorHtml(active.content, active.language) }} onMessage={e => { try { const msg=JSON.parse(e.nativeEvent.data); if(msg.type==="change"){setDraft(msg.content);setDirty(msg.content!==active.content);} else if(msg.type==="runResult"){setRunOutput(msg.output);setScreen("runner");} } catch {} }} javaScriptEnabled domStorageEnabled style={{ backgroundColor:C.bg }}/>
      </View>
      <Text style={{ color:C.muted, fontSize:11, paddingHorizontal:14, paddingBottom:8 }}>{dirty?"Cambios sin guardar":"Cambios guardados"} · {active.name}</Text>
    </View>}

    {screen==="projects" && <ScrollView contentContainerStyle={{ padding:16, gap:10 }}>
      <Text style={{ color:C.text, fontSize:21, fontWeight:"800" }}>Archivos locales</Text>
      <Text style={{ color:C.muted, marginBottom:8 }}>Guardados en esta instalación de la aplicación.</Text>
      {docs.map(d => <Pressable key={d.id} onPress={() => {setActiveId(d.id);setScreen("editor");}} style={{ padding:15, borderRadius:12, backgroundColor:C.panel, borderWidth:1, borderColor:C.border, flexDirection:"row", justifyContent:"space-between" }}>
        <View><Text style={{ color:C.text, fontWeight:"700" }}>{d.name}</Text><Text style={{ color:C.muted, fontSize:12, marginTop:5 }}>{languages.find(l=>l.id===d.language)?.label??d.language}</Text></View><Text style={{ color:C.accent }}>Abrir →</Text>
      </Pressable>)}
      <Pressable onPress={createFile} style={{ padding:15, borderRadius:10, borderStyle:"dashed", borderWidth:1, borderColor:C.accent, alignItems:"center" }}><Text style={{ color:C.accent, fontWeight:"700" }}>＋ Nuevo archivo</Text></Pressable>
    </ScrollView>}

    {screen==="languages" && <View style={{ flex:1, padding:14 }}>
      <Text style={{ color:C.text, fontSize:21, fontWeight:"800", marginBottom:10 }}>35 lenguajes y formatos</Text>
      <TextInput value={search} onChangeText={setSearch} placeholder="Buscar lenguaje…" placeholderTextColor={C.muted} style={{ color:C.text, backgroundColor:C.panel, borderWidth:1, borderColor:C.border, borderRadius:10, padding:12, marginBottom:10 }}/>
      <ScrollView>{filtered.map(l=><Pressable key={l.id} onPress={()=>{const next=docs.map(d=>d.id===activeId?{...d,language:l.id}:d);setDocs(next);setScreen("editor");AsyncStorage.setItem(KEY,JSON.stringify(next)).catch(()=>{});}} style={{ padding:13, borderBottomWidth:1, borderBottomColor:C.border, flexDirection:"row", justifyContent:"space-between" }}><Text style={{ color:C.text }}>{l.label}</Text><Text style={{ color:C.muted, fontSize:11 }}>{l.extensions.join(", ")}</Text></Pressable>)}</ScrollView>
    </View>}

    {screen==="github" && <ScrollView keyboardShouldPersistTaps="handled" contentContainerStyle={{ padding:15, gap:12 }}>
      <Text style={{ color:C.text, fontSize:22, fontWeight:"800" }}>GitHub</Text>
      <Text style={{ color:C.muted, lineHeight:20 }}>Cliente REST real. Requiere tu token de GitHub; no se conecta hasta que lo configures y la prueba responda correctamente. Para uso público, migra a GitHub App/OAuth con intercambio seguro.</Text>
      <Text style={{ color:C.muted, fontSize:12 }}>Token fine-grained · permiso Contents read/write para repositorios de prueba</Text>
      <TextInput value={tokenInput} onChangeText={setTokenInput} placeholder="github_pat_…" placeholderTextColor={C.muted} secureTextEntry autoCapitalize="none" autoCorrect={false} style={{ color:C.text, backgroundColor:C.panel, borderWidth:1, borderColor:C.border, borderRadius:10, padding:12 }}/>
      <View style={{ flexDirection:"row", gap:8 }}>
        <Pressable onPress={()=>withBusy(async()=>{await saveGitHubToken(tokenInput);setTokenInput("");setNotice("Token guardado en SecureStore. Pulsa Probar conexión.");})} style={{ flex:1, backgroundColor:C.accent, padding:12, borderRadius:9, alignItems:"center" }}><Text style={{ color:"#081126", fontWeight:"800" }}>Guardar token</Text></Pressable>
        <Pressable onPress={()=>withBusy(testGitHub)} style={{ flex:1, backgroundColor:C.green, padding:12, borderRadius:9, alignItems:"center" }}><Text style={{ color:"#062116", fontWeight:"800" }}>Probar conexión</Text></Pressable>
      </View>
      <Pressable onPress={()=>withBusy(removeToken)}><Text style={{ color:C.red, paddingVertical:5 }}>Eliminar token guardado</Text></Pressable>
      {user && <Text style={{ color:C.green, fontWeight:"700" }}>Conectado como @{user.login}</Text>}
      <Text style={{ color:C.text, fontSize:17, fontWeight:"800", marginTop:4 }}>Repositorios ({repos.length})</Text>
      {repos.map(r=><Pressable key={r.id} onPress={()=>chooseRepo(r)} style={{ padding:12, borderRadius:10, backgroundColor:selectedRepo?.id===r.id?C.panel2:C.panel, borderWidth:1, borderColor:selectedRepo?.id===r.id?C.accent:C.border }}>
        <Text style={{ color:C.text, fontWeight:"700" }}>{r.full_name}</Text><Text style={{ color:C.muted, fontSize:11, marginTop:3 }}>{r.private?"Privado":"Público"} · rama predeterminada: {r.default_branch}</Text>
      </Pressable>)}
      {selectedRepo && <>
        <Text style={{ color:C.text, fontSize:17, fontWeight:"800", marginTop:6 }}>{selectedRepo.full_name}</Text>
        <Text style={{ color:C.muted }}>Ramas</Text>
        <ScrollView horizontal style={{ maxHeight:42 }} contentContainerStyle={{ gap:6 }}>
          {branches.map(b=><Pressable key={b.name} onPress={()=>setBranch(b.name)} style={{ backgroundColor:branch===b.name?C.accent:C.panel2, padding:9, borderRadius:8 }}><Text style={{ color:branch===b.name?"#081126":C.text, fontSize:12 }}>{b.name}</Text></Pressable>)}
        </ScrollView>
        <TextInput value={remotePath} onChangeText={setRemotePath} placeholder="Ruta de archivo, ej. src/main.py" placeholderTextColor={C.muted} autoCapitalize="none" style={{ color:C.text, backgroundColor:C.panel, borderWidth:1, borderColor:C.border, borderRadius:9, padding:11 }}/>
        <Pressable onPress={()=>withBusy(loadRemoteFile)} style={{ backgroundColor:C.panel2, borderWidth:1, borderColor:C.border, padding:12, borderRadius:9, alignItems:"center" }}><Text style={{ color:C.text, fontWeight:"700" }}>Importar archivo remoto</Text></Pressable>
        <TextInput value={commitMessage} onChangeText={setCommitMessage} placeholder="Mensaje del commit" placeholderTextColor={C.muted} style={{ color:C.text, backgroundColor:C.panel, borderWidth:1, borderColor:C.border, borderRadius:9, padding:11 }}/>
        <Pressable onPress={()=>withBusy(pushCurrentFile)} style={{ backgroundColor:C.green, padding:12, borderRadius:9, alignItems:"center" }}><Text style={{ color:"#062116", fontWeight:"800" }}>Publicar archivo actual como commit</Text></Pressable>
        <Text style={{ color:C.muted, fontSize:12 }}>Comparar ramas</Text>
        <ScrollView horizontal contentContainerStyle={{ gap:6 }}>
          {branches.map(b=><Pressable key={`base-${b.name}`} onPress={()=>setBaseBranch(b.name)} style={{ padding:8, borderRadius:8, backgroundColor:baseBranch===b.name?C.accent:C.panel2 }}><Text style={{ color:baseBranch===b.name?"#081126":C.text, fontSize:11 }}>Base: {b.name}</Text></Pressable>)}
        </ScrollView>
        <ScrollView horizontal contentContainerStyle={{ gap:6 }}>
          {branches.map(b=><Pressable key={`head-${b.name}`} onPress={()=>setCompareBranch(b.name)} style={{ padding:8, borderRadius:8, backgroundColor:compareBranch===b.name?C.accent:C.panel2 }}><Text style={{ color:compareBranch===b.name?"#081126":C.text, fontSize:11 }}>Comparar: {b.name}</Text></Pressable>)}
        </ScrollView>
        <Pressable onPress={()=>withBusy(compareBranches)} style={{ padding:12, borderRadius:9, borderWidth:1, borderColor:C.accent, alignItems:"center" }}><Text style={{ color:C.accent, fontWeight:"800" }}>Comparar cambios</Text></Pressable>
        {!!compareText && <Text selectable style={{ color:C.text, backgroundColor:C.panel, padding:12, borderRadius:9, fontSize:11, lineHeight:17 }}>{compareText}</Text>}
      </>}
    </ScrollView>}

    {screen==="runner" && <ScrollView contentContainerStyle={{ padding:17, gap:12 }}>
      <Text style={{ color:C.text, fontSize:22, fontWeight:"800" }}>Ejecución</Text>
      <Text style={{ color:C.muted, lineHeight:20 }}>JavaScript local básico · Python y otros lenguajes aún no tienen intérprete configurado.</Text>
      <View style={{ padding:14, borderRadius:12, backgroundColor:C.panel, borderWidth:1, borderColor:C.border, gap:8 }}>
        <Text style={{ color:C.text, fontWeight:"800" }}>Salida del programa</Text>
        <Text selectable style={{ color:runOutput.startsWith("Ejecutando")?C.orange:C.green, lineHeight:21, fontFamily:"monospace" }}>{runOutput || "Abre un archivo JavaScript y pulsa ▶ Ejecutar."}</Text>
        <Pressable onPress={()=>{setRunOutput("");setScreen("editor");}} style={{ backgroundColor:C.panel2, padding:11, borderRadius:8, alignItems:"center" }}><Text style={{ color:C.text, fontWeight:"700" }}>Volver al editor</Text></Pressable>
      </View>
      <View style={{ padding:14, borderRadius:12, backgroundColor:C.panel, borderWidth:1, borderColor:C.border, gap:8 }}>
        <Text style={{ color:C.orange, fontWeight:"800" }}>Estado: JavaScript básico local</Text>
        <Text style={{ color:C.muted, lineHeight:21 }}>El botón Ejecutar funciona para JavaScript sencillo y muestra console.log, advertencias y errores. No es un entorno Node.js y no debe usarse para código no confiable. Python y otros lenguajes siguen pendientes de un motor dedicado.</Text>
      </View>
      <Text style={{ color:C.text, fontWeight:"800" }}>Plan técnico</Text>
      {["Local: JavaScript básico en el WebView; no equivale a Node.js ni a un sandbox de seguridad.","Remoto: API autenticada que envía tareas a contenedores efímeros.","Seguridad: tiempo máximo, límites CPU/memoria, sin secretos ni acceso libre a red.","Resultados: stdout, stderr, código de salida y duración."].map(x=><Text key={x} style={{ color:C.muted, lineHeight:21 }}>• {x}</Text>)}
      
    </ScrollView>}
  </SafeAreaView>;
}
function decodeBase64(input: string): string {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const clean = input.replace(/[^A-Za-z0-9+/=]/g, "");
  const bytes: number[] = [];
  for (let i=0; i<clean.length; i+=4) {
    const a=alphabet.indexOf(clean[i]), b=alphabet.indexOf(clean[i+1]);
    const c=clean[i+2]==="="?0:alphabet.indexOf(clean[i+2]);
    const d=clean[i+3]==="="?0:alphabet.indexOf(clean[i+3]);
    if(a<0||b<0) continue;
    bytes.push((a<<2)|(b>>4));
    if(clean[i+2]!=="="&&c>=0) bytes.push(((b&15)<<4)|(c>>2));
    if(clean[i+3]!=="="&&d>=0) bytes.push(((c&3)<<6)|d);
  }
  try { return new TextDecoder("utf-8").decode(new Uint8Array(bytes)); }
  catch { return bytes.map(b=>String.fromCharCode(b)).join(""); }
}