<template>
  <div class="container">
    <div v-if="!loading" class="fileList">
      <div v-for="file in files" :key="file.id" class="fileItem">
        <a :href="file.webUrl"><img :src="file.imageUrl"/></a>
        <a class="fileName" :href="file.webUrl" :title="file.name">{{ file.name }}</a>
        <span>{{ file.size | sizer }}</span>
      </div>
    </div>
    <div v-if="loading" class="lds-roller"><div></div><div></div><div></div><div></div><div></div><div></div><div></div><div></div></div>
  </div>
</template>

<script lang="ts">
import { Component, Prop, Vue } from 'vue-property-decorator';
import { Providers, prepScopes } from '@microsoft/mgt';

import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

@Component({
  filters: {
    sizer(value: any) {
      if (!value) {
        return '';
      }

      const size = value / 1000;

      if (size < 1000) {
        return Math.round(size) + 'KB';
      }

      return Math.round(size / 1000) + 'MB';
    },
  },
})
export default class RecentFileList extends Vue {
  private files: MicrosoftGraph.DriveItem[] = [];

  private loading = true;

  private async created() {
    if (Providers.globalProvider) {
      this.files = (await this.getFiles()).splice(0, 6);

      this.loading = false;
    }
  }

  private async getFiles(): Promise<MicrosoftGraph.DriveItem[]> {
    const client = Providers.globalProvider.graph.client;

    const uri = `/me/drive/recent`;

    const fileView = await client
      .api(uri)
      .middlewareOptions(prepScopes('files.read'))
      .get();

    if (fileView && fileView.value) {
      const files = fileView.value;
      //// files.sort((f1: MicrosoftGraph.DriveItem, f2: MicrosoftGraph.DriveItem) => f1.name.localeCompare(f2.name));

      for (const file of files) {
        const turi =
          `/drives/${file.remoteItem.parentReference.driveId}/items/${file.remoteItem.id}/thumbnails/0/medium`;

        try {
          const thumbnail = await client
            .api(turi)
            .middlewareOptions(prepScopes('files.read'))
            .get();

          file.imageUrl = thumbnail.url;
        } catch (err) {
          file.imageUrl = null; // TODO
        }
      }

      return files;
    }

    return [];
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped lang="scss">
  .container {
    display: flex;
    align-items: center;

    overflow-x: auto;

    height: 100%;
  }

  .fileList {
    padding-top: 16px;
    padding-left: 32px;
    
    display: flex;
    flex-direction: row;
    flex-wrap: nowrap;
    align-items: center;
  }

  .fileItem {
    margin-right: 53px;
    max-width: 120px;

    display: flex;
    flex-direction: column;

    img {
      display: block;
      height: 68px;
      width: auto;
      object-fit: contain;

      margin: 3px 6px 0px 0px;
    }

    span {
      font-style: normal;
      font-weight: 600;
      font-size: 10px;
      line-height: 13px;

      color: #767676;
    }
  }

  .fileName {
    overflow: hidden;
    white-space: nowrap;

    font-style: normal;
    font-weight: 600;
    font-size: 11.19px;
    line-height: 15px;

    color: #333333;
    text-decoration: none;

    text-overflow: ellipsis;

    &:hover {
      color: blue;
      text-decoration: underline;
    }
  }

  .lds-roller {
    display: inline-block;
    position: relative;
    width: 64px;
    height: 64px;

    left: 64px;
  }
  .lds-roller div {
    animation: lds-roller 1.2s cubic-bezier(0.5, 0, 0.5, 1) infinite;
    transform-origin: 32px 32px;
  }
  .lds-roller div:after {
    content: " ";
    display: block;
    position: absolute;
    width: 6px;
    height: 6px;
    border-radius: 50%;
    background: #D73B02;
    margin: -3px 0 0 -3px;
  }
  .lds-roller div:nth-child(1) {
    animation-delay: -0.036s;
  }
  .lds-roller div:nth-child(1):after {
    top: 50px;
    left: 50px;
  }
  .lds-roller div:nth-child(2) {
    animation-delay: -0.072s;
  }
  .lds-roller div:nth-child(2):after {
    top: 54px;
    left: 45px;
  }
  .lds-roller div:nth-child(3) {
    animation-delay: -0.108s;
  }
  .lds-roller div:nth-child(3):after {
    top: 57px;
    left: 39px;
  }
  .lds-roller div:nth-child(4) {
    animation-delay: -0.144s;
  }
  .lds-roller div:nth-child(4):after {
    top: 58px;
    left: 32px;
  }
  .lds-roller div:nth-child(5) {
    animation-delay: -0.18s;
  }
  .lds-roller div:nth-child(5):after {
    top: 57px;
    left: 25px;
  }
  .lds-roller div:nth-child(6) {
    animation-delay: -0.216s;
  }
  .lds-roller div:nth-child(6):after {
    top: 54px;
    left: 19px;
  }
  .lds-roller div:nth-child(7) {
    animation-delay: -0.252s;
  }
  .lds-roller div:nth-child(7):after {
    top: 50px;
    left: 14px;
  }
  .lds-roller div:nth-child(8) {
    animation-delay: -0.288s;
  }
  .lds-roller div:nth-child(8):after {
    top: 45px;
    left: 10px;
  }
  @keyframes lds-roller {
    0% {
      transform: rotate(0deg);
    }
    100% {
      transform: rotate(360deg);
    }
  }
</style>
