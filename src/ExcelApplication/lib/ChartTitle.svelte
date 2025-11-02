<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import ExcelBorder from "./ExcelBorder.svelte";
    import ExcelFont from "./ExcelFont.svelte";
    import {GetChart} from "../Scripts/GetChart";
    import UpdateFormatProperties from "../Scripts/UpdateFormatProperties";
    import UpdateGenericProperties from "../Scripts/UpdateGenericProperties";
    import { lang } from "../../Scripts/Language";
    import type { HelperType } from "../../Scripts/HelperType";
    import HelperDialogs from "../../lib/HelperDialogs.svelte";
    let helper = $state<HelperType | undefined>(undefined);

    const {data, spinner}: {data: Excel.Chart, spinner: HTMLDivElement} = $props();
</script>

{#if helper}
  <HelperDialogs helperType={helper} callback={() => (helper = undefined)}></HelperDialogs>
{/if}

<h2>{lang("Chart title")}:</h2>
<Card secondCard={true}>
    <h3>{lang("Position")}:</h3>
    <label class="flex hcenter gap">
        {lang("Top")}: <input type="number" bind:value={data.title.top} />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Left")}: <input type="number" bind:value={data.title.left} />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Text orientation")}: <input
            type="number"
            bind:value={data.title.textOrientation}
        />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Horizontal alignment")}: <div class="selectContainer"><select
            bind:value={data.title.horizontalAlignment}
        >
            {#each ["Center", "Left", "Right", "Justify", "Distributed"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label><br />
    <label class="flex hcenter gap">
        {lang("Vertical alignment")}: <div class="selectContainer"><select bind:value={data.title.verticalAlignment}>
            {#each ["Center", "Bottom", "Top", "Justify", "Distributed"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label><br />
    <label class="flex hcenter gap">
        {lang("Position")}: <div class="selectContainer"><select bind:value={data.title.position}>
            {#each ["Top", "Bottom", "Left", "Right"] as option}
            <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Font")}:</h3>
    <ExcelFont font={data.title.format.font}></ExcelFont>
</Card><br><Card secondCard={true}>
    <h3>{lang("Border")}:</h3>
    <ExcelBorder border={data.title.format.border}></ExcelBorder>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Other settings")}:</h3>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.title.showShadow}><span class="help" onclick={() => (helper = "ExcelTitleShadow")}>{lang("Add a shadow behind the title")}</span>
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.title.overlay}>{lang("Permit the title to overlay the chart")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.title.visible}>{lang("Make the title visible")}
    </label><br>
</Card><br>
<button onclick={() => {
    document.body.append(spinner);
    setTimeout(() => {
        Excel.run(async (ctx) => {
            const chart = await GetChart(ctx);
            chart.title.load({$all: true, format: {$all: true, border: {$all: true}, font: {$all: true}}});
            await ctx.sync();
            UpdateGenericProperties(data.title, chart.title, ["width", "height"]);
            await UpdateFormatProperties(data.title.format, chart.title.format, ctx);
            await ctx.sync();
            spinner.remove();
        })
    }, 1)
}}>{lang("Apply")}</button>