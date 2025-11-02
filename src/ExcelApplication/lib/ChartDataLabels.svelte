<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import { ExcelAvailableShapes, type HelperType } from "../../Scripts/HelperType";
    import ExcelBorder from "./ExcelBorder.svelte";
    import ExcelFont from "./ExcelFont.svelte";
    import {GetChart} from "../Scripts/GetChart";
    import UpdateFormatProperties from "../Scripts/UpdateFormatProperties";
    import { lang } from "../../Scripts/Language";
    import MissingProperties from "./MissingProperties.svelte";
    import HelperDialogs from "../../lib/HelperDialogs.svelte";

    const {
        dataLabels,
        isEmbedded,
        spinner
    }: {
        dataLabels: Excel.ChartDataLabel | Excel.ChartDataLabels;
        isEmbedded?: boolean;
        spinner: HTMLDivElement
    } = $props();
    let helper = $state<HelperType | undefined>(undefined);
</script>

{#if helper}
  <HelperDialogs helperType={helper} callback={() => (helper = undefined)}></HelperDialogs>
{/if}

{#if isEmbedded}
    <h4 class="help" onclick={() => (helper = "ExcelDataLabel")}>{lang("Data label")}:</h4>
{:else}
    <h2 class="help" onclick={() => (helper = "ExcelDataLabel")}>{lang("Data labels")}:</h2>
{/if}
<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={dataLabels.autoText} />
    {lang("Automatically generate the text inside the label")}
</label><br />
<label class="flex hcenter gap">
    {lang("Shape type")}: <div class="selectContainer"><select bind:value={dataLabels.geometricShapeType}>
        {#each ExcelAvailableShapes as shape}
            <option value={shape}>{shape}</option>
        {/each}
    </select></div>
</label><br />
<label class="flex hcenter gap">
    {lang("Separator string")}: <input type="text" bind:value={dataLabels.separator} />
</label><br />
<Card secondCard={true}>
    {#if isEmbedded}
        <h5>{lang("Position")}:</h5>
    {:else}
        <h3>{lang("Position")}:</h3>
    {/if}
    <label class="flex hcenter gap">
        {lang("Horizontal alignment")}: <div class="selectContainer"><select
            bind:value={dataLabels.horizontalAlignment}
        >
            {#each ["Center", "Left", "Right", "Justify", "Distributed"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label><br />
    <label class="flex hcenter gap">
        {lang("Vertical alignment")}: <div class="selectContainer"><select bind:value={dataLabels.verticalAlignment}>
            {#each ["Center", "Bottom", "Top", "Justify", "Distributed"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label><br />
    <label class="flex hcenter gap">
        {lang("Text rotation")}: <input
            type="number"
            bind:value={dataLabels.textOrientation}
        />
    </label><br />
    <label class="flex hcenter gap">
        {lang("Text position")}: <div class="selectContainer"><select bind:value={dataLabels.position}>
            {#each ["Invalid", "None", "Center", "InsideEnd", "InsideBase", "OutsideEnd", "Left", "Right", "Top", "Bottom", "BestFit", "Callout"] as option}
                <option value={option}>{lang(option)}</option>
            {/each}
        </select></div>
    </label>
</Card><br />
<Card secondCard={true}>
    {#if isEmbedded}
        <h5>{lang("Number format")}:</h5>
    {:else}
        <h3>{lang("Number format")}:</h3>
    {/if}
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={dataLabels.linkNumberFormat} />{lang("Link the number format to the one used in the axis")}
    </label><br />
    <label class="flex hcenter gap">
        {lang("Custom number format")}: <input
            type="text"
            bind:value={dataLabels.numberFormat}
        />
    </label>
</Card><br />
{#if Object.keys(dataLabels.format?.border ?? {}).length !== 0 }
<Card secondCard={true}>
    {#if isEmbedded}
        <h5>{lang("Border")}:</h5>
    {:else}
        <h3>{lang("Border")}:</h3>
    {/if}
    <ExcelBorder border={dataLabels.format.border}></ExcelBorder>
</Card><br>
{/if}
{#if Object.keys(dataLabels.format?.font ?? {}).length !== 0}
<Card secondCard={true}>
    {#if isEmbedded}
        <h5>{lang("Font")}:</h5>
    {:else}
        <h3>{lang("Font")}:</h3>
    {/if}
    <ExcelFont font={dataLabels.format.font}></ExcelFont>
</Card><br>
{/if}
{#if Object.keys(dataLabels.format?.border ?? {}).length === 0 || Object.keys(dataLabels.format?.font ?? {}).length === 0}
<MissingProperties textType={isEmbedded ? 5 : 3}></MissingProperties>
{/if}
<Card secondCard={true}>
    {#if isEmbedded}
        <h5>{lang("Advanced")}:</h5>
    {:else}
        <h3>{lang("Advanced")}:</h3>
    {/if}
    <p>{lang("Some charts display different things. Pick the ones that apply to your chart.")}</p>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={dataLabels.showBubbleSize} />{lang("Show bubble size")}
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={dataLabels.showCategoryName} />{lang("Show category name")}
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={dataLabels.showLegendKey} />{lang("Show legend key")}
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={dataLabels.showPercentage} />{lang("Show percentage")}
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={dataLabels.showSeriesName} />{lang("Show series name")}
    </label><br />
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={dataLabels.showValue} />{lang("Show value")}
    </label>
</Card>
{#if !isEmbedded}
<br>
<button onclick={() => {
    document.body.append(spinner);
    setTimeout(() => {
        Excel.run(async (ctx) => {
            const chart = await GetChart(ctx);
            chart.dataLabels.load({$all: true});
            await ctx.sync();
            let tryUpdatingFormatProperties = true;
            try {
                chart.dataLabels.format.load({$all: true, border: {$all: true}, font: {$all: true}});
                await ctx.sync();
            } catch(ex) {
                console.warn(ex);
                tryUpdatingFormatProperties = false;
            }
            for (const prop in dataLabels) {
                if (typeof dataLabels[prop as "separator"] !== "object" && dataLabels[prop as "separator"] !== chart.dataLabels[prop as "separator"]) chart.dataLabels[prop as "separator"] = dataLabels[prop as "separator"];
            }
            if (Object.keys(dataLabels.format?.border ?? {}).length !== 0) await UpdateFormatProperties(dataLabels.format, chart.dataLabels.format, ctx);
            await ctx.sync();
            spinner.remove();
        })
    }, 1);
}}>{lang("Apply")}</button>
{/if}
