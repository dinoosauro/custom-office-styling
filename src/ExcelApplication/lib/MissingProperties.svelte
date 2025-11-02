<script lang="ts">
    import Card from "../../lib/Card.svelte";
    import { lang } from "../../Scripts/Language";
    const {textType}: {textType?: 3 | 5} = $props(); 
</script>
<Card secondCard={true}>
    {#if textType === 5}
    <h5>{lang("Hidden properties")}:</h5>
    {:else}
    <h3>{lang("Hidden properties")}:</h3>
    {/if}
    <p>{lang("Since the element you're editing hasn't been added yet, you can't edit some of its properties. After you've added it, go back to the home by clicking on the extension logo on the top of the page to access them")}.</p>
</Card><br>