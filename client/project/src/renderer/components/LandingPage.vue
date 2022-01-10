<template>
  <div id="wrapper">
    <main>
      <div class="left-side">
        <system-information></system-information>
        <div class="postData" @click="getAllData">测试</div>
        <span> {{ titleInfo }}</span>
      </div>

      <div class="right-side">
        <div>
          <ul class="title" v-for="title in titleInfo" :key="title.name">
            {{
              title.showname
            }}
            <!-- 二级菜单 -->
            <div v-if="title.secList != null">
              <ul
                class="sectitle"
                v-for="sectitle in title.secList"
                :key="sectitle.name"
              >
                {{
                  sectitle.showname
                }}
                <input id="sectitle_des" v-model="sectitle.des" />
              </ul>
            </div>
            <!-- 一级菜单内容描述 -->
            <div v-if="title.des != undefined" id="title_des">
              <input id="title_des" v-model="title.des" />
            </div>
            <!-- 一级菜单合计 -->
            <div v-if="title.countcell != undefined" id="title_countcell">
              <input type="number" v-model="title.countcell" />
            </div>
          </ul>
        </div>
      </div>
    </main>
  </div>
</template>

<script>
import SystemInformation from "./LandingPage/SystemInformation";
const { ipcRenderer } = require("electron");

export default {
  name: "landing-page",
  components: { SystemInformation },
  data() {
    return {
      pageName: "LandingPage",
      titleInfo: [],
    };
  },
  watch: {
    // titleInfo: function () {
    //   console.log(this.titleInfo);
    // },
  },
  beforeCreate() {
    console.log("beforecreate:");
  },
  created() {
    console.log("create:");
    console.log(this.pageName);
    ipcRenderer.send("pageFinishInit", this.pageName);
  },
  beforeMount() {
    console.log("beforeMount:");
  },
  mounted() {
    console.log("mounted:");

    ipcRenderer.on("initTitle", (event, title, secTitles) => {
      this.titleInfo = this.handleTitle(title, secTitles);
    });
  },
  methods: {
    getTitleObj: function (item, type) {
      let obj = new Object({
        name: item.name,
        showname: item.showname,
      });
      if (type == "des") {
        obj["des"] = ""; //内容描述
      } else if (type == "countcell") {
        obj["countcell"] = 0; ///合计数值
      } else {
        obj["des"] = ""; //内容描述
        obj["countcell"] = 0; ///合计数值
      }
      return obj;
    },
    getSecTitleByTitleName: function (titleName, secTitles) {
      let result = null;
      for (const secTitleParentName in secTitles) {
        if (titleName == secTitleParentName) {
          if (result == null) {
            result = [];
          }
          result = secTitles[secTitleParentName];
        }
      }
      return result;
    },
    //处理json数据-二级菜单数据
    handleSecTitle: function (secTitleList) {
      if (secTitleList == null) {
        return null;
      }
      let result = [];
      for (const index in secTitleList) {
        let item = secTitleList[index];
        let secObj = this.getTitleObj(item, "des");
        result.push(secObj);
      }
      return result;
    },
    //处理json数据
    handleTitle: function (titles, secTitles) {
      let result = [];
      if (titles != null) {
        for (const index in titles) {
          let title_item = titles[index];

          let secTitleList = this.getSecTitleByTitleName(
            title_item.name,
            secTitles
          );
          let obj =
            secTitleList != null
              ? this.getTitleObj(title_item, "countcell")
              : title_item.addCountCell
              ? this.getTitleObj(title_item, "both")
              : title_item.countcell
              ? this.getTitleObj(title_item, "countcell")
              : this.getTitleObj(title_item, "des");
          let secList = this.handleSecTitle(secTitleList);
          obj["secList"] = secList;
          result.push(obj);
        }
      }
      return result;
    },
    getAllData: function () {
      let result = this.makeSaveJsonData(this.titleInfo);
      console.log(result);
      ipcRenderer.send("onSaveJson", result);
    },
    makeSaveJsonData: function (info) {
      let result = new Object();
      let obj = null;
      for (const index in info) {
        let item = info[index];
        if (item.name == "date") {
          result[item.des] = new Object();
          obj = result[item.des];
        } else if (item.name == "weekday") {
          continue;
        } else {
          let undefinedDes = item.des == undefined;
          let nullDes = item.des == "";
          let undefinedCountcell = item.countcell == undefined;
          let zeroCountcell = item.countcell == 0;
          let bSecList = item.secList != undefined && item.secList != null;
          if (!undefinedDes && !nullDes) {
            // 有描述-无合计
            if (undefinedCountcell) {
              console.log(item.name, item.des, obj);
              obj[item.name] = item.des;
            }
            // 有描述-有合计-无二级菜单
            else if (!bSecList) {
              obj[item.name] = new Object({
                des: item.des,
                countcell: parseInt(item.countcell),
              });
            }
          } else {
            //无描述-有二级菜单-有合计
            if (bSecList) {
              let secObj = new Object();
              let secList = item.secList;
              for (const key in secList) {
                const element = secList[key];
                if (element.des != undefined && element.des != "") {
                  secObj[element.name] = element.des;
                }
              }
              secObj["countcell"] = parseInt(item.countcell);
              obj[item.name] = secObj;
            } else {
              //无描述-无二级菜单-有合计 存在才记录
              if (!undefinedCountcell && !zeroCountcell) {
                obj[item.name] = new Object({
                  countcell: parseInt(item.countcell),
                });
              }
            }
          }
        }
      }
      return result;
    },
  },
};
</script>

<style>
@import url("https://fonts.googleapis.com/css?family=Source+Sans+Pro");

* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
}

body {
  font-family: "Source Sans Pro", sans-serif;
}

#wrapper {
  background: radial-gradient(
    ellipse at top left,
    rgba(255, 255, 255, 1) 40%,
    rgba(229, 229, 229, 0.9) 100%
  );
  height: 100vh;
  padding: 60px 80px;
  width: 100vw;
}

#logo {
  height: auto;
  margin-bottom: 20px;
  width: 420px;
}

main {
  display: flex;
  justify-content: space-between;
}

main > div {
  flex-basis: 50%;
}

.left-side {
  display: flex;
  flex-direction: column;
}

.title {
  color: #2c3e50;
  font-size: 14px;
  font-weight: bold;
  margin-bottom: 6px;
}
.sectitle {
  margin-left: 20px;
  margin-top: 6px;
}

.postData {
  width: 150px;
  height: 30px;
  line-height: 30px;
  text-align: center;
  border: 1px solid #ccc;
  text-align: center;
  cursor: pointer;
}
</style>
