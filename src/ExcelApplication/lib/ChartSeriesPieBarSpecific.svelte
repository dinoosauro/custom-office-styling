<script lang="ts">
    import HelperDialogs from "../../lib/HelperDialogs.svelte";
    import type { HelperType } from "../../Scripts/HelperType";
    import { lang } from "../../Scripts/Language";

    const { series, hideGap }: { series: Excel.ChartSeries, hideGap?: boolean } = $props();
    let helper = $state<HelperType | undefined>(undefined);

</script>

{#if helper}
  <HelperDialogs helperType={helper} callback={() => (helper = undefined)}></HelperDialogs>
{/if}


{#if !hideGap}
<label class="flex hcenter gap">
    <span class="help" onclick={() => (helper = "ExcelDistanceLength")}>{lang("Gap width")}:</span> <input type="number" bind:value={series.gapWidth} />
</label><br />
{/if}
<label class="flex hcenter gap">
    {lang("Second chart size")}: <input
        type="number"
        bind:value={series.secondPlotSize}
        min="5"
        max="200"
    />
</label><br>
<label class="flex hcenter gap">
    {lang("How the two sections should be split")}: <div class="selectContainer"><select bind:value={series.splitType}>
        <option value="SplitByPosition">{lang("Position")}</option>
        <option value="SplitByValue">{lang("Value")}</option>
        <option value="SplitByPercentValue">{lang("Percent value")}</option>
        <option value="SplitByCustomSplit">{lang("Custom split")}</option>
    </select></div>
</label><br>
<label class="flex hcenter gap">
    {lang("Treshold value for splitting the two sections")}: <input type="number" bind:value={series.splitValue}>
</label><br>