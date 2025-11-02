<script lang="ts">
    import { lang } from "../../Scripts/Language";
    const {font}: {font: Excel.ChartFont} = $props();
</script>

<label class="flex hcenter gap">
    {lang("Name")}: <input type="text" bind:value={font.name}>
</label><br>
<label class="flex hcenter gap">
    {lang("Color")}: <input type="color" bind:value={font.color}>
</label><br>
<label class="flex hcenter gap">
    {lang("Size")}: <input type="number" bind:value={font.size}>
</label><br>
<div class="flex hcenter gap" style="flex-wrap: wrap;">
    {#each [["bold", "Bold"], ["italic", "Italic"], ["underline", "Underline"]] as [fontStyle, fontText]}
        <label class="flex hcenter gap" style="gap: 5px">
            <input
                style="margin-right: 5px"
                type="checkbox"
                bind:checked={font[fontStyle as "bold"]}
            />
            {lang(fontText)}
        </label>
    {/each}
</div><br />
