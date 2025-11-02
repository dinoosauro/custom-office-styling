<script lang="ts">
    import { lang } from "../../Scripts/Language";

    const {line}: {line: Excel.ChartLineFormat} = $props();
</script>
<label class="flex hcenter gap">
    {lang("Color")}: <input type="color" bind:value={line.color}>
</label><br>
<label class="flex hcenter gap">
    {lang("Weight")}: <input type="number" bind:value={line.weight}>
</label><br>
<label class="flex hcenter gap">
    {lang("Style")}: <div class="selectContainer"><select bind:value={line.lineStyle}>
        {#each ["None" ,"Continuous" ,"Dash" ,"DashDot" ,"DashDotDot" ,"Dot" ,"Grey25" ,"Grey50" ,"Grey75" ,"Automatic" ,"RoundDot"] as option}
        <option value={option}>{option}</option>
        {/each}
    </select></div>
</label><br>
