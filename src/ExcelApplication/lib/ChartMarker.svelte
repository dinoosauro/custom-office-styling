<script lang="ts">
    import { lang } from "../../Scripts/Language";

    const {
        marker,
    }: {
        marker: {
            markerBackgroundColor: string;
            markerForegroundColor: string;
            markerSize: number;
            markerStyle: string;
        };
    } = $props();
</script>

<label class="flex hcenter gap">
    {lang("Foreground color")}: <input
        type="color"
        bind:value={marker.markerForegroundColor}
    />
</label><br />
<label class="flex hcenter gap">
    {lang("Background color")}: <input
        type="color"
        bind:value={marker.markerBackgroundColor}
    />
</label><br />
<label class="flex hcenter gap">
    {lang("Size")}: <input
        type="number"
        bind:value={marker.markerSize}
        min="2"
        max="72"
    />
</label><br />
<label class="flex hcenter gap">
    {lang("Style")}: <div class="selectContainer"> <select bind:value={marker.markerStyle}>
        {#each ["Invalid", "Automatic", "None", "Square", "Diamond", "Triangle", "X", "Star", "Dot", "Dash", "Circle", "Plus", "Picture"] as option}
            <option value={option}>{option}</option>
        {/each}
    </select></div>
</label>
