import axios from "axios";
import { cfg } from "../config"; // 👈 OJO: ../ porque estás dentro de scripts
import { graphGet } from "../graph";
import { getGraphToken } from "../auth";
async function main() {
    // 1) Intento v1.0 con includeHiddenLists
    try {
        const v1 = await graphGet(`/sites/${cfg.siteId}/lists?includeHiddenLists=true&$select=id,displayName`);
        console.log("=== LISTAS (v1.0 includeHiddenLists) ===");
        v1.value.forEach(l => console.log(l.displayName, "->", l.id));
        return;
    }
    catch (e) {
        console.log("v1.0 no devolvió ocultas, probando beta...");
    }
    // 2) Fallback beta
    const token = await getGraphToken();
    const { data } = await axios.get(`https://graph.microsoft.com/beta/sites/${cfg.siteId}/lists?includeHiddenLists=true&$select=id,displayName`, { headers: { Authorization: `Bearer ${token}` } });
    console.log("=== LISTAS (beta includeHiddenLists) ===");
    data.value.forEach((l) => console.log(l.displayName, "->", l.id));
}
main().catch(err => {
    console.error(err?.response?.data || err);
    process.exit(1);
});
