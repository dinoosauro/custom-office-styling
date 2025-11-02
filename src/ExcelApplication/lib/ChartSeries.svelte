<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import ChartDataLabels from "./ChartDataLabels.svelte";
    import ChartLineOptions from "./ChartLineOptions.svelte";
    import ChartMarker from "./ChartMarker.svelte";
    import ChartSeriesPieBarSpecific from "./ChartSeriesPieBarSpecific.svelte";
    import ExcelBorder from "./ExcelBorder.svelte";
    import {GetChart} from "../Scripts/GetChart";
    import UpdateFormatProperties from "../Scripts/UpdateFormatProperties";
    import UpdateGenericProperties from "../Scripts/UpdateGenericProperties";
    import { lang } from "../../Scripts/Language";
    import HelperDialogs from "../../lib/HelperDialogs.svelte";
    import type { HelperType } from "../../Scripts/HelperType";
    let helper = $state<HelperType | undefined>(undefined);
    const {data, spinner}: {data: Excel.Chart, spinner: HTMLDivElement} = $props();
    // @ts-ignore
    if (Array.isArray(data.series)) data.series = {items: data.series};
    for (const series of data.series.items) {
        // @ts-ignore
        if (Array.isArray(series.points)) series.points = {items: series.points}
    }
    let editItem = $state(1);
    let editPoint = $state(1);
    let gradientEdit = $state<"Maximum" | "Minimum" | "Midpoint">("Maximum");
</script>

{#if helper}
  <HelperDialogs helperType={helper} callback={() => (helper = undefined)}></HelperDialogs>
{/if}

<h2>{lang("Series")}:</h2>
<label class="flex hcenter gap">
    {lang("Edit series number")}: <input type="number" max={data.series.items.length} bind:value={editItem} min="1">
</label><br>
{#if data.series.items[editItem - 1].chartType === "Histogram" || data.series.items[editItem - 1].chartType === "Pareto"}
<Card secondCard={true}>
    <h3>{lang("Histogram interval options")}:</h3>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].binOptions.allowOverflow}>{lang("Allow overflow")}</label><br>
    <label class="flex hcenter gap">
        {lang("Overflow value")}: <input type="number" bind:value={data.series.items[editItem - 1].binOptions.overflowValue}>
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].binOptions.allowUnderflow}> {lang("Allow underflow")}
    </label><br>
    <label class="flex hcenter gap">
        {lang("Underflow value")}: <input type="number" bind:value={data.series.items[editItem - 1].binOptions.underflowValue}>
    </label><br>
    <label class="flex hcenter gap">
        {lang("Width")}: <input type="number" bind:value={data.series.items[editItem - 1].binOptions.width}>
    </label>
</Card><br>
{/if}
{#if data.series.items[editItem - 1].chartType === "Boxwhisker"}
<Card secondCard={true}>
    <h3>{lang("Boxwhisker chart options")}:</h3>
    <label class="flex hcenter gap">
        {lang("Quartile calculation")}: <div class="selectContainer"><select bind:value={data.series.items[editItem - 1].boxwhiskerOptions.quartileCalculation}>
            <option value="Inclusive">{lang("Inclusive")}</option>
            <option value="Exclusive">{lang("Exclusive")}</option>
        </select></div>
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].boxwhiskerOptions.showInnerPoints}>{lang("Show inner points")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].boxwhiskerOptions.showMeanLine}>{lang("Show mean line")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].boxwhiskerOptions.showMeanMarker}>{lang("Show mean marker")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].boxwhiskerOptions.showOutlierPoints}>{lang("Show outlier points")}
    </label>
</Card><br>
{/if}
{#if data.series.items[editItem - 1].chartType.toLowerCase().indexOf("bubble") !== -1}
<Card secondCard={true}>
    <h3>{lang("Bubble chart options")}:</h3>
    <label class="flex hcenter gap">
        {lang("Bubble size (scaling in percentage)")}: <input type="number" bind:value={data.series.items[0].bubbleScale}>
    </label>
</Card><br>
{/if} 
{#if data.series.items[editItem - 1].chartType.toLocaleLowerCase().indexOf("doughnut") !== -1 || data.series.items[editItem - 1].chartType.toLocaleLowerCase().indexOf("pie") !== -1}
<Card secondCard={true}>
    <h3>{lang("Chart-type specific options")}:</h3>
    {#if data.series.items[editItem - 1].chartType.toLocaleLowerCase().indexOf("doughnut") !== -1}
    <label class="flex hcenter gap">
        {lang("Doughnut hole size")}: <input type="number" bind:value={data.series.items[editItem - 1].doughnutHoleSize}>
    </label>
    {:else}
        <ChartSeriesPieBarSpecific series={data.series.items[editItem - 1]}></ChartSeriesPieBarSpecific>
    {/if}
    <label class="flex hcenter gap">
        <span class="help" onclick={() => (helper = "ExcelChartExplosion")}>{lang("Explosion")}:</span> <input type="number" bind:value={data.series.items[editItem - 1].explosion}>
    </label><br>
    <label class="flex hcenter gap">
        <span class="help" onclick={() => (helper = "ExcelChartAngle")}>{lang("Angle of the first slice")}:</span> <input type="number" bind:value={data.series.items[editItem - 1].firstSliceAngle}>
    </label>
</Card><br>
{/if}
{#if data.series.items[editItem - 1].chartType.toLowerCase().indexOf("bar") !== -1 || data.series.items[editItem - 1].chartType.toLowerCase().indexOf("column") !== -1}
<Card secondCard={true}>
    <h3>{lang("Chart-type specific options")}:</h3>
    <label class="flex hcenter gap">
        <span class="help" onclick={() => (helper = "ExcelDistanceLength")}>{lang("Gap width")}:</span> <input type="number" bind:value={data.series.items[editItem - 1].gapWidth}>
    </label><br>
    {#if data.series.items[editItem - 1].chartType.toLowerCase().indexOf("bar") !== -1 }
    <ChartSeriesPieBarSpecific series={data.series.items[editItem - 1]} hideGap={true}></ChartSeriesPieBarSpecific>
    {/if}
    <label class="flex hcenter gap">
        {lang("Overlap")}: <input type="number" bind:value={data.series.items[editItem - 1].overlap} min="-100" max="100">
    </label>
</Card><br>
{/if}
{#if data.series.items[editItem - 1].chartType === "RegionMap"}
<Card secondCard={true}>
    <h3>{lang("Region map specific settings")}:</h3>
    <label class="flex hcenter gap">
        {lang("Gradient styling")}: <div class="selectContainer"><select>
            <option value="TwoPhaseGroup">{lang("Two phase group")}</option>
            <option value="ThreePhaseGroup">{lang("Three phase group")}</option>
        </select></div>
    </label>
    <Card>
        <h4>{lang("Region levels")}:</h4>
        <label class="flex hcenter gap">
            {lang("Edit gradient options of the")}: <div class="selectContainer"><select bind:value={gradientEdit}>
                <option value="Maximum">{lang("Maximum")}</option>
                <option value="Midpoint">{lang("Midpoint")}</option>
                <option value="Minimum">{lang("Minimum")}</option>
            </select> </div>
        </label><br>
        <label class="flex hcenter gap">
            {lang("Color")}: <input type="color" bind:value={data.series.items[editItem - 1][`gradient${gradientEdit}Color`]}>
        </label><br>
        <label class="flex hcenter gap">
            {lang("Type")}: <div class="selectContainer"><select bind:value={data.series.items[editItem - 1][`gradient${gradientEdit}Type`]}>
                <option value="ExtremeValue">{lang("Extreme value")}</option>
                <option value="Number">{lang("Number")}</option>
                <option value="Percent">{lang("Percentage")}</option>
            </select></div>
        </label><br>
        <label class="flex hcenter gap">
            {lang("Maximum value")}: <input type="number" bind:value={data.series.items[editItem - 1][`gradient${gradientEdit}Value`]}>
        </label>
    </Card>
</Card><br>
{/if}
{#if data.series.items[editItem - 1].chartType === "Treemap"}
<Card secondCard={true}>
    <h3>{lang("Treemap-specific chart settings")}:</h3>
    <label class="flex hcenter gap">
        {lang("Parent label position")}: <div class="selectContainer"><select bind:value={data.series.items[editItem - 1].parentLabelStrategy}>
            <option value="None">None</option>
            <option value="Banner">Banner</option>
            <option value="Overlapping">Overlapping</option>
        </select></div>
    </label>
</Card><br>
{/if}
{#if data.series.items[editItem - 1].chartType === "Waterfall"}
<Card secondCard={true}>
    <h3>{lang("Waterfall-specific chart settings")}:</h3>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].showConnectorLines}>{lang("Show connector lines")}
    </label>
</Card><br>
{/if}
<Card secondCard={true}>
    <h3>{lang("Display settings")}:</h3>
    <label class="flex hcenter gap">
        {lang("Name")}: <input type="text" bind:value={data.series.items[editItem - 1].name} maxlength="255">
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].showShadow}>{lang("Add a shadow effect")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].varyByCategories}>
        {lang("Each item in the series should have a different color")}
    </label><br>
    <Card>
        <h4>{lang("Negative points")}:</h4>
        <label class="flex hcenter gap">
            {lang("Color for negative data points")}: <input type="color" bind:value={data.series.items[editItem - 1].invertColor}>
        </label><br>
        <label class="flex hcenter gap">
            <input type="checkbox" bind:checked={data.series.items[editItem - 1].invertIfNegative}>
            {lang("Invert the pattern if negative")}
        </label>
    </Card><br>
    <Card>
        <h4>{lang("Marker settings")}:</h4>
        <ChartMarker marker={data.series.items[editItem - 1]}></ChartMarker>
    </Card>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Format")}:</h3>
    <ChartLineOptions line={data.series.items[editItem - 1].format.line}></ChartLineOptions>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Single point settings")}:</h3>
    <p>{lang(`Note that you need to manually apply these settings with the "Apply" button of this card, and not with the one on the bottom of the page. You can change these settings for all the markers from the "Display settings" card above`)}.</p>
    <label class="flex hcenter gap">
        {lang("Edit point number")}: <input type="number" bind:value={editPoint} max={data.series.items[editPoint - 1].points.items.length}>
    </label><br>
    <Card>
        <h4>{lang("Marker")}:</h4>
        <ChartMarker marker={data.series.items[editItem - 1].points.items[editPoint - 1]}></ChartMarker>
    </Card><br>
    <Card>
        <h4>{lang("Border of the legend square")}:</h4>
        <ExcelBorder border={data.series.items[editItem - 1].points.items[editPoint - 1].format.border}></ExcelBorder>
    </Card><br>
    <Card>
        <ChartDataLabels spinner={spinner} isEmbedded={true} dataLabels={data.series.items[editItem - 1].points.items[editPoint - 1].dataLabel}></ChartDataLabels><br>
        <Card secondCard={true}>
            <h5>{lang("Position")}:</h5>
            <label class="flex hcenter gap">
                {lang("Top")}: <input type="number" bind:value={data.series.items[editItem - 1].points.items[editPoint - 1].dataLabel.top}>
            </label><br>
            <label class="flex hcenter gap">
                {lang("Left")}: <input type="number" bind:value={data.series.items[editItem - 1].points.items[editPoint - 1].dataLabel.left}>
            </label><br>
        </Card>    
    </Card><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].points.items[editPoint - 1].hasDataLabel}>{lang("Show data label")}
    </label><br>
    <button onclick={() => {
        document.body.append(spinner);
        setTimeout(() => {
            Excel.run(async (ctx) => {
                const chart = await GetChart(ctx);
                chart.series.load();
                await ctx.sync();
                chart.series.items[editItem - 1].points.load({$all: true, format: {border: {$all: true}}});
                await ctx.sync();
                try {
                    chart.series.items[editItem - 1].points.items[editPoint - 1].dataLabel.load({$all: true, format: {$all: true, border: {$all: true}, font: {$all: true}}});
                    await ctx.sync();
                } catch(ex) {
                    console.warn(ex);
                }
                UpdateGenericProperties(data.series.items[editItem - 1].points.items[editPoint - 1], chart.series.items[editItem - 1].points.items[editPoint - 1]);
                if (typeof data.series.items[editItem - 1].points.items[editPoint - 1].format !== "undefined") await UpdateFormatProperties(data.series.items[editItem - 1].points.items[editPoint - 1].format, chart.series.items[editItem - 1].points.items[editPoint - 1].format, ctx);
                if (Object.keys(data.series.items[editItem - 1].points.items[editPoint - 1].dataLabel?.format ?? {}).length !== 0) {
                    UpdateGenericProperties(data.series.items[editItem - 1].points.items[editPoint - 1].dataLabel, chart.series.items[editItem - 1].points.items[editPoint - 1].dataLabel, ["height", "width"]);
                    await UpdateFormatProperties(data.series.items[editItem - 1].points.items[editItem - 1].dataLabel.format, chart.series.items[editItem - 1].points.items[editPoint - 1].dataLabel.format, ctx);
                }
                await ctx.sync();
                spinner.remove();
            })
        }, 1)

    }}>{lang("Apply")}</button>
</Card><br>
<Card secondCard={true}>
    <h3>{lang("Advanced")}:</h3>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].filtered}>{lang("Filtered chart")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].hasDataLabels}>{lang("Add a label with the data")}
    </label><br>
    <label class="flex hcenter gap">
        {lang("Plot order")}: <input type="number" bind:value={data.series.items[editItem - 1].plotOrder}>
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].showLeaderLines}>
        {lang("Show leader lines")}
    </label><br>
    <label class="flex hcenter gap">
        <input type="checkbox" bind:checked={data.series.items[editItem - 1].smooth}>
        {lang("Make the series smooth")}
    </label>
</Card><br>
<button onclick={() => {
    document.body.append(spinner);
    setTimeout(() => {
        Excel.run(async (ctx) => {
            const chart = await GetChart(ctx);
            chart.series.load();
            await ctx.sync();
            chart.series.items[editItem - 1].load({$all: true, format: {$all: true, line: {$all: true}}});
            await ctx.sync();
            UpdateGenericProperties(data.series.items[editItem - 1], chart.series.items[editItem - 1]);
            await UpdateFormatProperties(data.series.items[editItem - 1].format, chart.series.items[editItem - 1].format, ctx);
            await ctx.sync();
            spinner.remove();
        })
    }, 1)
}}>{lang("Apply")}</button>