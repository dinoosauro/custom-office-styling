<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import ExcelBorder from "./ExcelBorder.svelte";
    import ExcelFont from "./ExcelFont.svelte";
    import UpdateFormatProperties from "../Scripts/UpdateFormatProperties";
    import { lang } from "../../Scripts/Language";
    const { data, spinner }: { data: Excel.Chart, spinner: HTMLDivElement } = $props();
</script>

<h2>{lang("Chart data")}:</h2>
<Card secondCard={true}>
    <h3>{lang("Border")}:</h3>
    <ExcelBorder border={data.format.border}></ExcelBorder>
</Card><br />
<Card secondCard={true}>
    <h3>{lang("Font")}:</h3>
    <ExcelFont font={data.format.font}></ExcelFont>
</Card><br />
<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={data.format.roundedCorners} />{lang("Rounded corners")}
</label><br />
<button
    onclick={() => {
        document.body.append(spinner);
        setTimeout(() => {
            Excel.run(async (ctx) => {
                let temptempData = ctx.workbook.worksheets
                    .getActiveWorksheet()
                    .charts.load();
                await ctx.sync();
                let chart = temptempData.items[0];
                chart.format.load({$all: true, font: {$all: true}, border: {$all: true}});
                await ctx.sync();
                await UpdateFormatProperties(data.format, chart.format, ctx);
                await ctx.sync();
                spinner.remove();
            });
        }, 1)
    }}>{lang("Apply")}</button
>
