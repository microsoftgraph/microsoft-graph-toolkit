<template>
  <div id="app"
       :class="{ 'loggedOut' : !signedIn, 'loggedIn' : signedIn }">
    <mgt-msal-provider 
      client-id="a974dfa0-9f57-49b9-95db-90f04ce2111a"
      scopes="user.read,people.read,user.readbasic.all,contacts.read,calendars.read,files.read,group.read.all,tasks.readwrite">
    </mgt-msal-provider>

    <header v-if="!signedIn">
      <h1>My Organizer</h1>
      <p>Sign-in to get started organizing your work and life.</p>
      <mgt-login v-pre>
        <template data-type="signed-out">
          <div>Sign In</div>
        </template>
      </mgt-login>
    </header>

    <logged-in-page v-if="signedIn"/>
  </div>
</template>

<script lang="ts">
import { Component, Vue } from 'vue-property-decorator';
import { MgtMsalProvider, MgtLogin, ProviderState, Providers } from '@microsoft/mgt';
import LoggedInPage from './components/LoggedInPage.vue';

@Component({
  components: {
    LoggedInPage,
  },
})
export default class App extends Vue {
  private signedIn = false;
  private browserOpen = false;

  private async created() {
    const provider = Providers.globalProvider;
    if (provider) {
      this.loggedIn(provider.state);
      provider.onStateChanged(this.loggedIn);
    }

    Providers.onProviderUpdated(() => {
      const provider2 = Providers.globalProvider;
      if (provider2) {
        this.loggedIn(provider2.state);
        provider2.onStateChanged(this.loggedIn);
      }
    });
  }

  private loggedIn(state: any) {
    if (state === ProviderState.SignedIn) {
      this.signedIn = true;
    }
  }

  private loggedOut() {
    this.signedIn = false;
  }
}
</script>

<style lang="scss">
  body {
    margin: 0px;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
  }

  #app {
    &.loggedOut {
      background: url('/background.svg');
      background-size: 100% 100%;
      height: 100vh;
    }
  }

  header {
    background-color: white;
    border: 1px solid #C8C6C4;
    box-sizing: border-box;
    box-shadow: 0px -4px 56px rgba(10, 38, 66, 0.05), 0px 4px 16px rgba(0, 0, 0, 0.07);
    
    width: 664px;
    height: 458px;

    margin: auto;

    position: relative;
    top: 50%;
    transform: perspective(1px) translateY(-50%);

    font-style: normal;
    text-align: center;

    h1 {
      font-size: 56px;
      font-weight: 350;
      line-height: 100%;/* or 56px */

      margin: 72px 0px 0px 0px;

      color: #323130;
    }

    p {
      font-size: 16px;
      line-height: 24px;/* or 150% */

      margin: 64px 0px 0px 0px;

      color: #323130;
    }

    mgt-login {    
      margin: 64px auto 0px auto;

      width: 160px;
      --width: 160px;
      height: 50px;
      --height: 50px;

      --color: white;
      --background-color: #0078D4;

      div {
        width: 100%;
      }
    }
  }

  #fileBrowser {
    position: absolute;

    top: 32px;
    left: 32px;

    padding: 32px;

    z-index: 9999;

    background: white;
    box-shadow: 0px -4px 56px rgba(10, 38, 66, 0.05), 0px 4px 16px rgba(0, 0, 0, 0.07);
  }
</style>
