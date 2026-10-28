Office.onReady(() => {});

const GRAPH = "https://graph.microsoft.com/v1.0";
const TOGGLE_YEARS = ["2024"];
const archivePath = (year) => `/me/drive/root:/MesArchives/${year}:/children`;
const storageKey = (year) => `archive${year}`;

function getToken() {
    return Office.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true });
}

async function graph(path, token, options = {}) {
    const response = await fetch(`${GRAPH}${path}`, {
        ...options,
        headers: { Authorization: "Bearer " + token, ...(options.headers || {}) }
    });
    if (!response.ok) {
        throw new Error(`Graph ${response.status} on ${path}`);
    }
    return response.status === 204 ? null : response.json();
}

function toBase64(text) {
    return btoa(unescape(encodeURIComponent(text)));
}

function readState(year) {
    try { return localStorage.getItem(storageKey(year)); } catch (e) { return null; }
}

function writeState(year, value) {
    try {
        if (value) { localStorage.setItem(storageKey(year), value); }
        else { localStorage.removeItem(storageKey(year)); }
    } catch (e) { console.error(e); }
}

// Red/green toggle: mounted => "unmount" button enabled, "mount" disabled; unmounted => reverse.
async function setMounted(year, mounted) {
    if (!TOGGLE_YEARS.includes(year) || !Office.ribbon || !Office.ribbon.requestUpdate) { return; }
    try {
        await Office.ribbon.requestUpdate({
            tabs: [{
                id: "TabDefault",
                groups: [{
                    id: "MesArchivesGroup",
                    controls: [
                        { id: `archive${year}_mount`, enabled: !mounted },
                        { id: `archive${year}_unmount`, enabled: mounted }
                    ]
                }]
            }]
        });
    } catch (e) {
        console.error("Ribbon update failed:", e);
    }
}

async function mount(year, event) {
    try {
        const token = await getToken();
        const files = await graph(archivePath(year), token);
        const emls = (files.value || []).filter((f) => /\.eml$/i.test(f.name) && f["@microsoft.graph.downloadUrl"]);
        const folder = await graph("/me/mailFolders", token, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ displayName: `Archives ${year}` })
        });
        writeState(year, folder.id);
        for (const file of emls) {
            try {
                const download = await fetch(file["@microsoft.graph.downloadUrl"]);
                if (!download.ok) { throw new Error(`Download ${download.status}`); }
                const mime = await download.text();
                await graph(`/me/mailFolders/${folder.id}/messages`, token, {
                    method: "POST",
                    headers: { "Content-Type": "text/plain" },
                    body: toBase64(mime)
                });
            } catch (e) {
                console.error(`Import failed for ${file.name}:`, e);
            }
        }
        await setMounted(year, true);
    } catch (error) {
        console.error(`Mount ${year} failed:`, error);
    }
    event.completed();
}

async function unmount(year, event) {
    try {
        const id = readState(year);
        if (id) {
            const token = await getToken();
            await graph(`/me/mailFolders/${id}`, token, { method: "DELETE" });
        }
        writeState(year, null);
        await setMounted(year, false);
    } catch (error) {
        console.error(`Unmount ${year} failed:`, error);
    }
    event.completed();
}

function openModop(event) {
    window.open('https://dellasiegaexternal2.github.io/MesArchivesAddin/ModOp_Mes_ARCHIVES.pdf', '_blank');
    event.completed({ allowEvent: true });
}

["2021", "2022", "2023", "2024"].forEach((year) => {
    Office.actions.associate(`mount${year}`, (event) => mount(year, event));
    Office.actions.associate(`unmount${year}`, (event) => unmount(year, event));
});
Office.actions.associate("openModop", openModop);
