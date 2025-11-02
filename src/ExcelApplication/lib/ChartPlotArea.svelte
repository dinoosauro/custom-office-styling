<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import ExcelBorder from "./ExcelBorder.svelte";
    import {GetChart} from "../Scripts/GetChart";
    import UpdateFormatProperties from "../Scripts/UpdateFormatProperties";
    import UpdateGenericProperties from "../Scripts/UpdateGenericProperties";
    import { lang } from "../../Scripts/Language";

    const {data, spinner }: {data: Excel.Chart, spinner: HTMLDivElement} = $props();
    let editInside = $state(false);
</script>

<h2>{lang("Chart area")}:</h2>
<label class="flex hcenter gap">
    {lang("Position")}: <div class="selectContainer"><select bind:value={data.plotArea.position}>
        <option value="Automatic">{lang("Automatic")}</option>
        <option value="Custom">{lang("Custom")}</option>
    </select></div>
</label><br>
<Card secondCard={true}>
    <h3>{lang("Chart area border")}:</h3>
    <ExcelBorder border={data.plotArea.format.border}></ExcelBorder>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Plot position in the chart")}:</h3>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={editInside}>{lang("Edit inside properties")}
    </label><br>
    <label class="flex hcenter gap">
        {lang("Top")}: <input type="number" bind:value={data.plotArea[editInside ? "insideTop" : "top"]}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Left")}: <input type="number" bind:value={data.plotArea[editInside ? "insideLeft" : "left"]}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Height")}: <input type="number" bind:value={data.plotArea[editInside ? "insideHeight" : "height"]}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Width")}: <input type="number" bind:value={data.plotArea[editInside ? "insideWidth" : "width"]}>
    </label><br>
</Card><br>
<button onclick={() => {
    document.body.append(spinner);
    setTimeout(() => {
        Excel.run(async (ctx) => {
            const chart = await GetChart(ctx);
            chart.plotArea.load({$all: true, format: {border: {$all: true}}});
            await ctx.sync();
            UpdateGenericProperties(data.plotArea, chart.plotArea);
            if (chart.plotArea.format.border.color !== null || data.plotArea.format.border.weight !== 0) await UpdateFormatProperties(data.plotArea.format, chart.plotArea.format, ctx);
            await ctx.sync();
            spinner.remove();
        })
    }, 1)
}}>{lang("Apply")}</button>