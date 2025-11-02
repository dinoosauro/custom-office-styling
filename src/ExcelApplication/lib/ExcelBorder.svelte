<script lang="ts">
    import { lang } from "../../Scripts/Language";

    const {border}: {border: Excel.ChartBorder} = $props();
    
</script>

<label class="flex hcenter gap">
    {lang("Color")}: <input type="color" bind:value={border.color}>
</label><br>
<label class="flex hcenter gap">
    {lang("Border type")}: <div class="selectContainer"><select bind:value={border.lineStyle}>
        {#each ["Continuous","Dash","DashDot","DashDotDot","Dot","Grey25","Grey50","Grey75","Automatic","RoundDot"] as option}
            <option value={option}>{option}</option>
        {/each}
    </select></div>
</label><br>
<label class="flex hcenter gap">
    {lang("Weight")}: <input type="number" bind:value={border.weight}>
</label>