<script lang="ts">
    import { lang } from "../../Scripts/Language";
import { GetChart } from "../Scripts/GetChart";
    let {width, height, urlCallback, spinner}: {width: number, height: number, urlCallback: (link: string) => void, spinner: HTMLDivElement} = $props();
    let fittingMode = "Fit";
</script>

<h2>{lang("Export chart image")}:</h2>
<label class="flex hcenter gap">
    {lang("Width")}: <input type="number" min="1" step="1" bind:value={width}>
</label><br>
<label class="flex hcenter gap">
    {lang("Height")}: <input type="number" min="1" step="1" bind:value={height}>
</label><br>
<label class="flex hcenter gap">
    {lang("Fitting mode")}: <div class="selectContainer"><select bind:value={fittingMode}>
        <option value="Fit">{lang("Fit")}</option>
        <option value="FitAndCenter">{lang("Fit and center")}</option>
        <option value="Fill">{lang("Fill")}</option>
    </select></div>
</label><br>
<button onclick={() => {
    document.body.append(spinner);
    setTimeout(() => {
        Excel.run(async (ctx) => {
            const chart = await GetChart(ctx);
            chart.title.load();
            await ctx.sync();
            const img = chart.getImage(width, height, fittingMode as "Fit");
            await ctx.sync();
            const url = new URL(window.location.href);
            url.pathname = `${url.pathname.substring(0, url.pathname.lastIndexOf("/"))}/downloader.html`;
            url.hash = new URLSearchParams({
                base64: "1",
                data: img.value,
                name: `${chart.title.text || chart.name || chart.id}.png`,
            }).toString();
            if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
                Office.context.ui.openBrowserWindow(url.toString());
            } else urlCallback(url.toString()); // Fallback: a dialog will be displayed with the link to download it.     
            spinner.remove();
        })
    }, 1)
}}>{lang("Export")}</button>