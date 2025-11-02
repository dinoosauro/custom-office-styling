<script lang="ts">
    import { lang } from "../../Scripts/Language";
import ChartLineOptions from "./ChartLineOptions.svelte";

    const {gridlineFormat}: {
        gridlineFormat: Excel.ChartGridlines
    } = $props();
</script>

{#if typeof gridlineFormat.format?.line !== "undefined" && Object.keys(gridlineFormat.format?.line).length !== 0}
<ChartLineOptions line={gridlineFormat.format?.line}></ChartLineOptions>
{/if}
<label class="flex hcenter gap">
    <input type="checkbox" bind:checked={gridlineFormat.visible}>{lang("Is visible")}
</label>