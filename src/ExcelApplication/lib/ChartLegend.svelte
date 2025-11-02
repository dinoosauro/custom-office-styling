<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import ChartFormat from "./ChartFormat.svelte";
    import ExcelBorder from "./ExcelBorder.svelte";
    import ExcelFont from "./ExcelFont.svelte";
    import {GetChart} from "../Scripts/GetChart";
    import UpdateFormatProperties from "../Scripts/UpdateFormatProperties";
    import { lang } from "../../Scripts/Language";
    import type { HelperType } from "../../Scripts/HelperType";
    import HelperDialogs from "../../lib/HelperDialogs.svelte";
    let helper = $state<HelperType | undefined>(undefined);
    const { data, spinner }: { data: Excel.Chart, spinner: HTMLDivElement } = $props();
    let selectedLegend = 1;
    let selectedLegendVisibility = true;
</script>

<h2>{lang("Legend")}:</h2>
<Card secondCard={true}>
    <h3>{lang("Single legend entry")}:</h3>
        <label class="flex hcenter gap">
            {lang("Edit legend number")}:
            <input type="number" min="1" bind:value={selectedLegend}>
        </label><br>
        <label class="flex hcenter gap">
            <input type="checkbox" bind:checked={selectedLegendVisibility}>{lang("Make it visible")}
        </label><br>
    <button onclick={() => {
        document.body.append(spinner);
        setTimeout(() => {
            Excel.run(async (ctx) => {
                let chart = await GetChart(ctx);
                const entries = chart.legend.legendEntries.load();
                await ctx.sync();
                entries.items[selectedLegend - 1].visible = selectedLegendVisibility;
                spinner.remove();
            })
        }, 1)
    }}>{lang("Apply")}</button>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("General settings")}:</h3>
    <label class="flex hcenter gap">
        {lang("Legend height")}: <input type="number" bind:value={data.legend.height} />
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.legend.overlay} /> {lang("Overlap the legend with the main chart")}
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.legend.showShadow} /> <span class="help" onclick={() => (helper = "ExcelAxisLegendShadow")}>{lang("Show a shadow in the legend background")}</span>
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.legend.visible} />{lang("Make the legend visible")}
    </label><br>
</Card><br />
<Card secondCard={true}>
    <h3>{lang("Legend position")}:</h3>
        <label class="flex hcenter gap">
        {lang("Legend position")}: <div class="selectContainer"> <select bind:value={data.legend.position}>
            {#each ["Top","Bottom","Left","Right","Corner","Custom"] as option}
            <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Top")}: <input type="number" bind:value={data.legend.top}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Left")}: <input type="number" bind:value={data.legend.left}>
    </label>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Font")}:</h3>
    <ExcelFont font={data.legend.format.font}></ExcelFont>
</Card><br />
<Card secondCard={true}>
    <h3>{lang("Border")}:</h3>
    <ExcelBorder border={data.legend.format.border}></ExcelBorder>
</Card><br />
<button
    onclick={() => {
        Excel.run(async (ctx) => {
            let temptempData = ctx.workbook.worksheets
                .getActiveWorksheet()
                .charts.load();
            await ctx.sync();
            let chart = temptempData.items[0];
            const legend = chart.legend.load({
                $all: true,
                format: {
                    $all: true,
                    border: { $all: true },
                    font: { $all: true },
                },
            });
            await ctx.sync();
            // We'll avoid putting top and left properties if the legend position has been changed
            const skipTopLeft = data.legend.position !== legend.position;
            for (const prop in data.legend) {
                if ((prop === "top" || prop === "left") && skipTopLeft) continue;
                if (typeof data.legend[prop as "overlay"] !== "object") {
                    if (
                        data.legend[prop as "overlay"] !==
                        legend[prop as "overlay"]
                    )
                        legend[prop as "overlay"] =
                            data.legend[prop as "overlay"];
                }
            }
            await UpdateFormatProperties(data.legend.format, chart.legend.format, ctx);
            await ctx.sync();
        });
    }}>{lang("Apply")}</button
>
{#if helper}
  <HelperDialogs helperType={helper} callback={() => (helper = undefined)}></HelperDialogs>
{/if}

