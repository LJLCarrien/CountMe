<template>
  <div id="wrapper">
    <main>
      <div class="left-side">
        <system-information></system-information>
        <div class="postData" @click="getAllData">
          <span>保存数据</span>
        </div>
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
              <input 
                v-if="title.name === 'date'" 
                type="date" 
                v-model="title.des" 
              />
              <input v-else id="title_des" v-model="title.des" />
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
    titleInfo: {
      handler: function () {
        this.updateWeekday();
      },
      deep: true,
    },
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
      this.$nextTick(() => {
        this.updateWeekday();
      });
    });
  },
  methods: {
    updateWeekday: function () {
      const dateItem = this.titleInfo.find(item => item.name === 'date');
      const weekdayItem = this.titleInfo.find(item => item.name === 'weekday');
      if (dateItem && weekdayItem && dateItem.des) {
        const weekDays = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
        weekdayItem.des = weekDays[new Date(dateItem.des).getDay()];
      }
    },
    getTitleObj: function (item, type) {
      let obj = new Object({
        name: item.name,
        showname: item.showname,
      });
      if (type == "des") {
        if (item.name === "date") {
          obj["des"] = new Date().toISOString().split('T')[0];
        } else if (item.name === "weekday") {
          const weekDays = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];
          obj["des"] = weekDays[new Date().getDay()];
        } else {
          obj["des"] = "";
        }
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
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
}

#wrapper {
  background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
  height: 100vh;
  padding: 40px 60px;
  width: 100vw;
  overflow: auto;
}

#logo {
  height: auto;
  margin-bottom: 20px;
  width: 420px;
}

main {
  display: flex;
  justify-content: space-between;
  gap: 40px;
}

main > div {
  flex-basis: 50%;
}

.left-side {
  display: flex;
  flex-direction: column;
  background: white;
  border-radius: 16px;
  padding: 24px;
  box-shadow: 0 10px 40px rgba(0, 0, 0, 0.1);
}

.right-side {
  background: white;
  border-radius: 16px;
  padding: 24px;
  box-shadow: 0 10px 40px rgba(0, 0, 0, 0.1);
}

.title {
  color: #2c3e50;
  font-size: 16px;
  font-weight: 600;
  margin-bottom: 8px;
  padding: 12px 16px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  border-radius: 8px;
  margin: 8px 0;
}

.sectitle {
  margin-left: 24px;
  margin-top: 8px;
  padding: 8px 12px;
  background: #f8f9fa;
  border-radius: 6px;
  border-left: 3px solid #667eea;
}

.postData {
  width: 180px;
  height: 44px;
  line-height: 44px;
  text-align: center;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  border: none;
  border-radius: 22px;
  cursor: pointer;
  font-size: 16px;
  font-weight: 500;
  margin-top: 20px;
  transition: all 0.3s ease;
  box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
}

.postData:hover {
  transform: translateY(-2px);
  box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
}

.postData:active {
  transform: translateY(0);
}

#title_des input,
#sectitle_des,
.title_countcell input {
  width: 100%;
  padding: 10px 14px;
  margin-top: 6px;
  border: 2px solid #e1e5eb;
  border-radius: 8px;
  font-size: 14px;
  transition: all 0.3s ease;
  outline: none;
}

#title_des input:focus,
#sectitle_des:focus,
.title_countcell input:focus {
  border-color: #667eea;
  box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
}

input[type="number"] {
  width: 120px;
  padding: 8px 12px;
  border: 2px solid #e1e5eb;
  border-radius: 8px;
  font-size: 14px;
  outline: none;
  transition: all 0.3s ease;
}

input[type="number"]:focus {
  border-color: #667eea;
  box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
}

#title_des,
.title_countcell {
  margin-top: 8px;
  padding: 8px 12px;
  background: #f8f9fa;
  border-radius: 6px;
}

.right-side > div {
  max-height: calc(100vh - 160px);
  overflow-y: auto;
}

.right-side::-webkit-scrollbar {
  width: 8px;
}

.right-side::-webkit-scrollbar-track {
  background: #f1f1f1;
  border-radius: 4px;
}

.right-side::-webkit-scrollbar-thumb {
  background: #c1c1c1;
  border-radius: 4px;
}

.right-side::-webkit-scrollbar-thumb:hover {
  background: #a1a1a1;
}
</style>
