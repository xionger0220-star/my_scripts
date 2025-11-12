import pandas as pd
import os
import re
from datetime import datetime
import requests


class ActorTestRange:
    def __init__(self,branch_name,project_path):
        self.project_path = project_path  #项目路径
        self.branch_name = branch_name   #分支名
        self.download_file()  #下载excel
        self.point_path = "./stats.details.xlsx"
        self.excel_path = f"{self.project_path}design/Excel"    #配置表路径
        #加载表格
        self.sheet_dict = self.get_sheet_dict()
        self.gameObject = self.get_config("GameObject")
        self.storageComponent = self.get_config("StorageComponent")
        self.drop = self.get_config("Drop")
        self.dispayTemplate = self.get_config("DispayTemplate")
        self.produce = self.get_config("Produce")
        self.produceList = self.get_config("ProduceList")
        self.produceComponent = self.get_config("ProduceComponent")
        self.burnComponent = self.get_config("BurnComponent")
        self.gatherComponent = self.get_config("GatherComponent")
        self.deathComponent = self.get_config("DeathComponent")
        self.hurtComponent = self.get_config("HurtComponent")
        self.useComponent = self.get_config("UseComponent")
        self.itemUse = self.get_config("ItemUse")
        self.taskReward = self.get_config("TaskReward")
        self.marketComponent = self.get_config("MarketComponent")
        self.marketType = self.get_config("MarketType")
        self.marketGoods = self.get_config("MarketGoods")
        self.gridSpawnInstance = self.get_config("GridSpawnInstance")
        self.pointCloudGroup = self.get_config("PointCloudGroup")
        self.pointCloud = self.get_config("PointCloud")
        self.task = self.get_config("Task")
        self.bornIdentity = self.get_config("BornIdentity")
        #获取地图数据
        self.map_point_id_dict = self.get_24KMap_point_id()  # 24K点云
        self.map_poi_id_list = self.get_24KMap_poi()  # 24K poi
        #获取投放
        self.map_item_dict = self.get_24kMap_id()   #24k投放
        self.poiStorage_item_dict = self.get_poiStorageItem_id()  #poi仓储投放
        self.produce_item_dict = self.get_produce_id()   #制造投放
        self.task_item_dict = self.get_task_reward_id()   #任务奖励投放
        self.task_private_item_dict = self.get_task_private_item_id()  #任务私有交互物
        self.character_dict = self.get_character_id()  #出生表投放
        self.market_item_dict = self.get_market_id()   #商店售卖
        self.map_animal_dict = self.get_24KMap_animal_id()  #24K投放的动物
        self.map_monster_dict = self.get_gridType_monster_id()   #宫格价值表投放的怪物
        self.poi_monster_dict = self.get_poiTrigger_monster_id()  #poi trigger投放的怪物
        self.chrPoint_monster_dict = self.get_chrPoint_monster_id()  #虚点刷出的怪物
        self.all_item_dict = self.get_all_item_dict()  #收集所有交互物
        #获取机制投放
        self.storage_drop_item_dict = self.get_storage_drop_id()  # 仓储机制投放
        self.display_item_dict = self.get_dispay_id()  #换装表投放
        self.hurt_drop_item_dict = self.get_hurt_drop_id()  #受伤机制投放
        self.burn_drop_item_dict = self.get_burn_drop_id()  #燃烧机制投放
        self.death_drop_item_dict = self.get_death_drop_id()  #死亡机制投放
        self.gather_drop_item_dict = self.get_gather_drop_id()  #采集机制投放
        self.use_drop_item_dict = self.get_use_drop_id() #使用机制投放

    # 获取sheet与表名的映射关系字典
    def get_sheet_dict(self):
        excel_path = self.excel_path
        sheet_dict = {}
        # 遍历目录下所有文件
        for filename in os.listdir(excel_path):
            # 仅处理Excel文件
            if filename.endswith(('.xlsx', '.xls')):
                file_path = os.path.join(excel_path, filename)
                try:
                    # 读取Excel文件的所有sheet名称
                    with pd.ExcelFile(file_path) as xls:
                        sheet_names = xls.sheet_names
                    # 更新字典
                    for sheet in sheet_names:
                        if sheet in sheet_dict:
                            if filename not in sheet_dict[sheet]:
                                sheet_dict[sheet].append(filename)
                        else:
                            sheet_dict[sheet] = [filename]
                except Exception as e:
                    print(f"处理文件 {filename} 时出错: {str(e)}")
                    continue
        return sheet_dict

    #将同sheet名的数据合并
    def get_config(self,sheet_name):
        excel_path = self.excel_path
        excel_files = self.sheet_dict[sheet_name]
        dataframes = []
        for excel_file in excel_files:
            file_path = os.path.join(excel_path, excel_file)
            # 读取指定sheet
            try:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=1).iloc[2:]
                dataframes.append(df)
            except Exception as e:
                print(f"{excel_file} sheet {sheet_name}: {str(e)}")
        if not dataframes:
            return pd.DataFrame()
        return pd.concat(dataframes, ignore_index=True)

    #传入id，机制名，返回机制id
    def get_object_mechanism_id(self,config_id,mechanism_name):
        filter = self.gameObject[(self.gameObject.object_id == int(config_id)) | (self.gameObject.object_id == str(config_id))]
        if len(filter) != 0:
            s = filter.object_mechanism_group.values[0]
            if s == s:
                pattern = r'\b' + re.escape(mechanism_name) + r':(\d+)'
                match = re.search(pattern, s)
                if match:
                    return int(match.group(1))
                else:
                    return f"交互物表没配{mechanism_name}机制"
            else:
                return f"交互物表没配{mechanism_name}机制"
        else:
            return f"交互物表找不到此id"

    #传入掉落id,返回掉落的道具列表
    def get_drop_item_list(self,drop_id):
        filter = self.drop[((self.drop.drop_id == int(drop_id)) | (self.drop.drop_id == str(drop_id))) & (self.drop.drop_num > 0) & (self.drop.order_weight_prob > 0)]
        drop_item_id_list = filter.drop_item_id.tolist()  # 掉落物品ID
        drop_drop_id_list = filter.drop_drop_id.tolist()  # 嵌套掉落ID
        drop_drop_id_list = [int(i) for i in drop_drop_id_list if i == i]
        drop_drop_id_str_list = [str(i) for i in drop_drop_id_list]
        drop_item_id_list += self.drop[(self.drop.drop_id.isin(drop_drop_id_list) | self.drop.drop_id.isin(drop_drop_id_str_list)) & (self.drop.drop_num > 0) & (self.drop.order_weight_prob > 0)].drop_item_id.tolist()
        drop_item_id_list = [int(i) for i in drop_item_id_list if i == i]
        return drop_item_id_list

    #合并字典
    def merge_dict(self,d1,d2):
        merged = d1.copy()  # 避免修改原字典
        for key, value in d2.items():
            if key not in merged:
                merged[key] = value  # 合并列表
            else:
                for des in value:
                    if des not in merged[key]:
                        merged[key] += [des]  #新增
        return merged

    #获取地图数据文件里的点云ID
    def get_mapData_id(self,path,suffixes,pattern):
        point_list = []
        for file_name in os.listdir(path):
            file_path = os.path.join(path, file_name)
            if file_name.endswith((suffixes)):
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    target_ids = re.findall(pattern, content)
                    for i in target_ids:
                        point_list.append(int(i))
        return list(set(point_list))

    #获取24K投放动物id
    def get_24KMap_animal_id(self):
        path = f"{self.project_path}design\SOCMapData\DGM01\PointCloud\Anim_Real\Txt"
        suffixes = '.txt'
        pattern = r'TypeID\s*:\s*(\d+)'
        animal_id_list = self.get_mapData_id(path,suffixes,pattern)
        monster_dict = {}
        for animal_id in animal_id_list:
            monster_dict[animal_id] = ["24K投放"]
        return monster_dict

    #获取虚点投放的怪物id
    def get_chrPoint_monster_id(self):
        path = f"{self.project_path}design\SOCMapData\DGM01\ChrPointGroup\Txt"
        suffixes = '.txt'
        pattern = r'TypeID\s*:\s*(\d+)'
        chrPoint_id_list = self.get_mapData_id(path, suffixes, pattern)
        monster_dict = {}
        for chrPoint_id in chrPoint_id_list:
            filter = self.pointCloudGroup[(self.pointCloudGroup.group_point_id == int(chrPoint_id)) | (self.pointCloudGroup.group_point_id == str(chrPoint_id))]
            monster_id_list = filter.point_cloud_id.tolist()
            monster_id_list = list(set(monster_id_list))
            for monster_id in monster_id_list:
                if int(monster_id) not in monster_dict:
                    monster_dict[int(monster_id)] = [f"{chrPoint_id}虚点投放"]
                elif f"{chrPoint_id}虚点投放" not in monster_dict[int(monster_id)]:
                    monster_dict[int(monster_id)] += [f"{chrPoint_id}虚点投放"]
        return monster_dict

    #获取宫格价值表投放怪物id
    def get_gridType_monster_id(self):
        monster_dict = {}
        group_point_id = self.gridSpawnInstance.group_point_id.tolist()
        group_point_id = list(set(group_point_id))
        group_point_id = [int(i) for i in group_point_id]
        group_point_id_str = [str(i) for i in group_point_id]
        filter = self.pointCloudGroup[self.pointCloudGroup.group_point_id.isin(group_point_id) | self.pointCloudGroup.group_point_id.isin(group_point_id_str)]
        if len(filter) != 0:
            monster_id_list = filter.point_cloud_id.tolist()
            monster_id_list = list(set(monster_id_list))
            for i in monster_id_list:
                monster_dict[int(i)] = ["宫格价值表投放"]
        return monster_dict

    #获取POI投放怪物id
    def get_poiTrigger_monster_id(self):
        file_path = f"{self.project_path}design\SOCMapData\DGM01\EnvPointGroup\Txt"
        poi_monster_dict = {}
        for poi_id in self.map_poi_id_list:
            json_path = f"{file_path}/{poi_id}_trigger.json"
            try:
                with open(json_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    target_ids = re.findall(r'"ConfigID":\s*(\d+)', content)
                    target_ids = list(set(target_ids))
                    for point_group_id in target_ids:
                        monster_id_list = self.pointCloudGroup[(self.pointCloudGroup.group_point_id == int(point_group_id)) | (self.pointCloudGroup.group_point_id == str(point_group_id))].point_cloud_id.tolist()
                        for monster_id in monster_id_list:
                            if int(monster_id) not in poi_monster_dict:
                                poi_monster_dict[int(monster_id)] = [f"POI-{poi_id}投放"]
                            elif f"POI-{poi_id}投放" not in poi_monster_dict[int(monster_id)]:
                                poi_monster_dict[int(monster_id)] += [f"POI-{poi_id}投放"]
            except FileNotFoundError:
                print(f"{json_path}找不到")
        return poi_monster_dict


    #获取24K里的点云id
    def get_24KMap_point_id(self):
        path = f"{self.project_path}design\SOCMapData\DGM01\PointCloud\Common\Txt"
        point_id_dict = {}
        for file_name in os.listdir(path):
            file_path = os.path.join(path, file_name)
            if file_name.endswith(('.txt')):
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    target_ids = re.findall(r'TypeID\s*:\s*(\d+)', content)
                    for i in target_ids:
                        point_id_dict[int(i)] = ["24K投放"]
        return point_id_dict

    #获取24K里的poi
    def get_24KMap_poi(self):
        poi_list = []
        for point_id,_ in self.map_point_id_dict.items():
            filter = self.pointCloud[(self.pointCloud.point_cloud_id == int(point_id)) | (self.pointCloud.point_cloud_id == str(point_id))]
            if (len(filter) != 0) and (filter.type.values[0] == "POI"):
                poi_list.append(int(point_id))
        poi_list = list(set(poi_list))
        return poi_list

    #获取24K投放表的id
    def get_24kMap_id(self):
        df = pd.read_excel(self.point_path, sheet_name="ConfigMetrics", header=0)
        map_id_list = df["配置 ID"].tolist()
        map_id_dict = {}
        for i in map_id_list:
            if int(i) not in map_id_dict:
                map_id_dict[int(i)] = ["24K投放"]
        return map_id_dict

    #获取POI仓储投放的id
    def get_poiStorageItem_id(self):
        df = pd.read_excel(self.point_path, sheet_name="POIRewardItems", header=0)
        poiStorageItem_id_list = df["配置 ID"].tolist()
        poiStorageItem_id_dict = {}
        for i in poiStorageItem_id_list:
            if int(i) not in poiStorageItem_id_dict:
                poiStorageItem_id_dict[int(i)] = ["POI仓储投放"]
        return poiStorageItem_id_dict

    #获取制造表可制造的id
    def get_produce_id(self):
        produce_id_list = self.produceList[self.produceList.is_produce == "YES"].produce_id.tolist()
        produce_id_list = [int(i) for i in produce_id_list]
        produce_id_str_list = [str(i) for i in produce_id_list]
        produce_id_list = self.produce[self.produce.produce_id.isin(produce_id_list) | self.produce.produce_id.isin(produce_id_str_list)].produce_reward.tolist()
        produce_id_dict = {}
        for i in produce_id_list:
            item_id = i.split(":")[0]
            if int(item_id) not in produce_id_dict:
                produce_id_dict[int(item_id)] = ["可制造"]
        return produce_id_dict

    #获取换装模板表投放的id
    def get_dispay_id(self):
        display_id_dict = {}
        for item_id, item_path in self.all_item_dict.items():
            storage_id = self.get_object_mechanism_id(item_id, "storage")
            if str(storage_id).isdigit():
                filter = self.storageComponent[(self.storageComponent.entity_component_id == int(storage_id)) | (self.storageComponent.entity_component_id == str(storage_id))]
                if len(filter) != 0:
                    display_id = filter.display_template_id.values[0]
                    if display_id == display_id:
                        filter = self.dispayTemplate[(self.dispayTemplate.display_template_id == int(display_id)) | (self.dispayTemplate.display_template_id == str(display_id))]
                        if len(filter) != 0:
                            display_id_list = filter.object_display_id.tolist()
                            id_list = []
                            for i in display_id_list:
                                id_list += str(i).split(";")
                            id_list = [int(i) for i in id_list]
                            for display_item_id in id_list:
                                if int(display_item_id) not in display_id_dict:
                                    display_id_dict[int(display_item_id)] = [f"{item_id}换装掉落"]
                                elif f"{item_id}换装掉落" not in display_id_dict[int(display_item_id)]:
                                    display_id_dict[int(display_item_id)] += [f"{item_id}换装掉落"]
        return display_id_dict

    #获取仓储机制投放的id列表
    def get_storage_drop_id(self):
        # 独立仓储掉落
        storage_drop_dict = {}
        for item_id,item_path in self.all_item_dict.items():
            storage_id = self.get_object_mechanism_id(item_id, "storage")
            if str(storage_id).isdigit():
                filter = self.storageComponent[(self.storageComponent.entity_component_id == int(storage_id)) | (self.storageComponent.entity_component_id == str(storage_id))]
                if len(filter) != 0:
                    drop_id = filter.personal_drop_id.values[0]
                    if drop_id == drop_id:
                        drop_item_id_list = self.get_drop_item_list(int(drop_id))
                        for drop_item_id in drop_item_id_list:
                            if int(drop_item_id) not in storage_drop_dict:
                                storage_drop_dict[int(drop_item_id)] = [f"{item_id}仓储掉落"]
                            elif f"{item_id}仓储掉落" not in storage_drop_dict[int(drop_item_id)]:
                                storage_drop_dict[int(drop_item_id)] += [f"{item_id}仓储掉落"]
        return storage_drop_dict

    #获取受伤机制掉落的id列表
    def get_hurt_drop_id(self):
        # 受伤机制掉落
        drop_dict = {}
        for item_id,item_path in self.all_item_dict.items():
            mechanism_id = self.get_object_mechanism_id(item_id, "hurt")
            if str(mechanism_id).isdigit():
                filter = self.hurtComponent[(self.hurtComponent.entity_component_id == int(mechanism_id)) | (self.hurtComponent.entity_component_id == str(mechanism_id))]
                if len(filter) != 0:
                    drop_id = filter.health_drop_group.values[0]
                    if drop_id == drop_id:
                        drop_item_id_list = self.get_drop_item_list(int(drop_id))
                        for drop_item_id in drop_item_id_list:
                            if int(drop_item_id) not in drop_dict:
                                drop_dict[int(drop_item_id)] = [f"{item_id}受伤掉落"]
                            elif f"{item_id}受伤掉落" not in drop_dict[int(drop_item_id)]:
                                drop_dict[int(drop_item_id)] += [f"{item_id}受伤掉落"]
        return drop_dict

    #获取燃烧机制掉落的id列表
    def get_burn_drop_id(self):
        # 燃烧机制掉落
        drop_dict = {}
        for item_id,item_path in self.all_item_dict.items():
            mechanism_id = self.get_object_mechanism_id(item_id, "burn")
            if str(mechanism_id).isdigit():
                filter = self.burnComponent[(self.burnComponent.entity_component_id == int(mechanism_id)) | (self.burnComponent.entity_component_id == str(mechanism_id))]
                if len(filter) != 0:
                    drop_id = filter.public_drop_id.values[0]
                    if drop_id == drop_id:
                        drop_item_id_list = self.get_drop_item_list(int(drop_id))
                        for drop_item_id in drop_item_id_list:
                            if int(drop_item_id) not in drop_dict:
                                drop_dict[int(drop_item_id)] = [f"{item_id}燃烧掉落"]
                            elif f"{item_id}燃烧掉落" not in drop_dict[int(drop_item_id)]:
                                drop_dict[int(drop_item_id)] += [f"{item_id}燃烧掉落"]
        return drop_dict

    #获取死亡机制掉落的id列表
    def get_death_drop_id(self):
        # 死亡机制掉落
        drop_dict = {}
        for item_id,item_path in self.all_item_dict.items():
            mechanism_id = self.get_object_mechanism_id(item_id, "death")
            if str(mechanism_id).isdigit():
                filter = self.deathComponent[(self.deathComponent.entity_component_id == int(mechanism_id)) | (self.deathComponent.entity_component_id == str(mechanism_id))]
                if len(filter) != 0:
                    drop_id = filter.personal_drop_id.values[0]
                    if drop_id == drop_id:
                        drop_item_id_list = self.get_drop_item_list(int(drop_id))
                        for drop_item_id in drop_item_id_list:
                            if int(drop_item_id) not in drop_dict:
                                drop_dict[int(drop_item_id)] = [f"{item_id}死亡掉落"]
                            elif f"{item_id}死亡掉落" not in drop_dict[int(drop_item_id)]:
                                drop_dict[int(drop_item_id)] += [f"{item_id}死亡掉落"]
        return drop_dict

    #获取采集机制掉落的id列表
    def get_gather_drop_id(self):
        # 采集机制掉落
        drop_dict = {}
        for item_id,item_path in self.all_item_dict.items():
            mechanism_id = self.get_object_mechanism_id(item_id, "gather")
            if str(mechanism_id).isdigit():
                filter = self.gatherComponent[(self.gatherComponent.entity_component_id == int(mechanism_id)) | (self.gatherComponent.entity_component_id == str(mechanism_id))]
                if len(filter) != 0:
                    drop_id = filter.personal_drop_id.values[0]
                    if drop_id == drop_id:
                        drop_item_id_list = self.get_drop_item_list(int(drop_id))
                        for drop_item_id in drop_item_id_list:
                            if int(drop_item_id) not in drop_dict:
                                drop_dict[int(drop_item_id)] = [f"{item_id}采集掉落"]
                            elif f"{item_id}采集掉落" not in drop_dict[int(drop_item_id)]:
                                drop_dict[int(drop_item_id)] += [f"{item_id}采集掉落"]
        return drop_dict

    #获取使用机制掉落的id列表
    def get_use_drop_id(self):
        # 使用机制掉落
        drop_dict = {}
        for item_id,item_path in self.all_item_dict.items():
            mechanism_id = self.get_object_mechanism_id(item_id, "use")
            if str(mechanism_id).isdigit():
                filter = self.useComponent[(self.useComponent.entity_component_id == int(mechanism_id)) | (self.useComponent.entity_component_id == str(mechanism_id))]
                if len(filter) != 0:
                    item_use_list = filter.item_use_id.tolist()
                    item_use_id_list = []
                    for s in item_use_list:
                        if s == s:
                            item_use_id_list += str(s).split(";")
                    item_use_id_list = [int(i) for i in item_use_id_list]
                    item_use_id_list = list(set(item_use_id_list))
                    item_use_id_list_str = [str(i) for i in item_use_id_list]
                    filter = self.itemUse[self.itemUse.item_use_id.isin(item_use_id_list) | self.itemUse.item_use_id.isin(item_use_id_list_str)]
                    #使用后获得固定道具
                    use_get_item_list = filter.get_item.tolist()
                    for part in use_get_item_list:
                        if part == part:
                            match = re.findall(r"(\d+):",part)
                            if match:
                                for use_get_item_id in match:
                                    if int(use_get_item_id) not in drop_dict:
                                        drop_dict[int(use_get_item_id)] = [f"{item_id}使用掉落"]
                                    elif f"{item_id}使用掉落" not in drop_dict[int(use_get_item_id)]:
                                        drop_dict[int(use_get_item_id)] += [f"{item_id}使用掉落"]
                    #使用后获得随机道具
                    drop_id_list = filter.drop_id.tolist()
                    drop_id_list = [int(i) for i in drop_id_list if i == i]
                    drop_id_list = list(set(drop_id_list))
                    for drop_id in drop_id_list:
                        drop_item_id_list = self.get_drop_item_list(int(drop_id))
                        for drop_item_id in drop_item_id_list:
                            if int(drop_item_id) not in drop_dict:
                                drop_dict[int(drop_item_id)] = [f"{item_id}使用掉落"]
                            elif f"{item_id}使用掉落" not in drop_dict[int(drop_item_id)]:
                                drop_dict[int(drop_item_id)] += [f"{item_id}使用掉落"]
        return drop_dict

    #出生表投放
    def get_character_id(self):
        character_dict = {}
        character_id_list = self.bornIdentity[self.bornIdentity['scence_id'].str.contains("2",na=False)].character_id.tolist()
        for item_id in character_id_list:
            if int(item_id) not in character_dict:
                character_dict[int(item_id)] = ["出生表投放"]
        return character_dict


    #获取任务奖励的id列表
    def get_task_reward_id(self):
        task_reward_dict = {}
        task_reward_list = self.taskReward[self.taskReward.is_blocked == "NO"].task_reward.tolist()
        task_reward_list = [i for i in task_reward_list if i==i]
        for i in task_reward_list:
            match = re.findall('(\d+):', i)
            if match:
                for item_id in match:
                    if int(item_id) not in task_reward_dict:
                        task_reward_dict[int(item_id)] = ["任务奖励"]
        return task_reward_dict


    #获取任务创建的私有交互物
    def get_task_private_item_id(self):
        task_private_item_dict = {}
        blocked_task_id_list = self.taskReward[self.taskReward.is_blocked == "YES"].task_id.tolist()
        blocked_task_id_list.append(1)
        df = self.task[~self.task.task_id.isin(blocked_task_id_list)]
        task_create_entity = df.task_create_entity.tolist()
        for i in task_create_entity:
            if i == i:
                for j in i.split(";"):
                    item_id = j.split(":")[0]
                    if int(item_id) not in task_private_item_dict:
                        task_private_item_dict[int(item_id)] = ["任务私有交互物"]
        task_create_entity_coordinate = df.task_create_entity_coordinate.tolist()
        for i in task_create_entity_coordinate:
            if i == i:
                for j in i.split(";"):
                    item_id = j.split(",")[2]
                    if int(item_id) not in task_private_item_dict:
                        task_private_item_dict[int(item_id)] = ["任务私有交互物"]
        task_create_entity_pointcloud = df.task_create_entity_pointcloud.tolist()
        for i in task_create_entity_pointcloud:
            if i == i:
                for j in i.split(";"):
                    item_id = j.split(",")[9]
                    if int(item_id) not in task_private_item_dict:
                        task_private_item_dict[int(item_id)] = ["任务私有交互物"]
        return task_private_item_dict


    #获取商店售卖的id列表
    def get_market_id(self):
        market_dict = {}
        for item_id,item_path in self.map_item_dict.items():
            mechanism_id = self.get_object_mechanism_id(item_id, "market")
            if str(mechanism_id).isdigit():
                filter = self.marketComponent[(self.marketComponent.entity_component_id == int(mechanism_id)) | (self.marketComponent.entity_component_id == str(mechanism_id))]
                if len(filter) != 0:
                    market_id_group = filter.market_id_group.values[0]
                    market_id_list = [int(i) for i in market_id_group.split(";")]
                    market_id_list_str = [str(i) for i in market_id_list]
                    filter = self.marketType[self.marketType.market_id.isin(market_id_list) | self.marketType.market_id.isin(market_id_list_str)]
                    if len(filter) != 0:
                        market_inventory_id_group = filter.market_inventory_id.tolist()
                        market_inventory_id_list = []
                        for s in market_inventory_id_group:
                            market_inventory_id_list += s.split(";")
                        market_inventory_id_list = [int(i) for i in market_inventory_id_list]
                        market_inventory_id_list = list(set(market_inventory_id_list))
                        market_inventory_id_list_str = [str(i) for i in market_inventory_id_list]
                        filter = self.marketGoods[self.marketGoods.market_inventory_id.isin(market_inventory_id_list) | self.marketGoods.market_inventory_id.isin(market_inventory_id_list_str)]
                        if len(filter) != 0:
                            goods_item = filter.goods_item.tolist()
                            for goods in goods_item:
                                goods_id = int(goods.split(":")[0])
                                if goods_id not in market_dict:
                                    market_dict[goods_id] = [f"{item_id}商店售卖"]
                                elif f"{item_id}商店售卖" not in market_dict[goods_id]:
                                    market_dict[goods_id] += [f"{item_id}商店售卖"]
        return market_dict

    #获取交互机制
    def get_mechanism_list(self,mechanism_str):
        mechanism_list = []
        button_def = {
            "EnergyOpen":"打开电源",
            "EnergyClose":"关闭电源",
            "Energy":"能源管理",
        }
        if mechanism_str == mechanism_str:
            mechanism_dict = {}
            for i in mechanism_str.split(";"):
                mechanism_dict[i.split(":")[0]] = int(i.split(":")[1])
            if "take_in" in mechanism_dict:
                pass
        return mechanism_list

    def get_all_item_dict(self):
        id_dict = self.map_item_dict      #24K地图投放
        id_dict = self.merge_dict(id_dict, self.poiStorage_item_dict)   #poi仓储投放
        id_dict = self.merge_dict(id_dict, self.produce_item_dict)  # 制造投放
        id_dict = self.merge_dict(id_dict, self.character_dict)   #角色出生自带
        id_dict = self.merge_dict(id_dict, self.task_item_dict)  #任务奖励投放
        id_dict = self.merge_dict(id_dict, self.task_private_item_dict)   #任务私有交互物
        id_dict = self.merge_dict(id_dict, self.market_item_dict)  #商店售卖
        id_dict = self.merge_dict(id_dict, self.map_animal_dict)  #24k 动物投放
        id_dict = self.merge_dict(id_dict, self.map_monster_dict) #24k 怪物投放
        id_dict = self.merge_dict(id_dict, self.poi_monster_dict)  #poi 怪物投放
        id_dict = self.merge_dict(id_dict, self.chrPoint_monster_dict)  # 虚点 怪物投放
        return id_dict

    #总投放
    def all_to_excel(self):
        #机制投放
        id_dict = self.merge_dict(self.all_item_dict, self.storage_drop_item_dict)  # 仓储机制投放
        id_dict = self.merge_dict(id_dict, self.display_item_dict)  #换装表投放
        id_dict = self.merge_dict(id_dict, self.hurt_drop_item_dict)  #受伤机制投放
        id_dict = self.merge_dict(id_dict, self.burn_drop_item_dict)  # 燃烧机制投放
        id_dict = self.merge_dict(id_dict, self.death_drop_item_dict)  # 死亡机制投放
        id_dict = self.merge_dict(id_dict, self.gather_drop_item_dict)  # 采集机制投放
        id_dict = self.merge_dict(id_dict, self.use_drop_item_dict)  # 使用机制投放
        data = []
        for item_id, delivery_method in id_dict.items():
            filter = self.gameObject[(self.gameObject.object_id == int(item_id)) | (self.gameObject.object_id == str(item_id))]
            if len(filter) != 0:
                name = filter.object_name.values[0]
                mechanism_list = self.get_mechanism_list(filter.object_mechanism_group.values[0])
                data.append([item_id, name, delivery_method, mechanism_list])
        df = pd.DataFrame(data, columns=["交互物id", "名字", "投放方式", "交互物机制"])
        timestamp = datetime.now().strftime("%Y%m%d%H%M")
        path = f"D:/交互物全量投放_{timestamp}.xlsx"
        df.to_excel(path, index=False)

    #输入分支名：master or stable
    def download_file(self):
        branch_dict = {
            "master" : "http://10.1.8.136:9010/gm/world/wd/loc_data/map_data_cooked/stats.details.xlsx",
            "stable" : "http://10.1.7.114:9010/gm/world/wd/loc_data/map_data_cooked/stats.details.xlsx"
        }
        url = branch_dict[self.branch_name]
        with requests.get(url, stream=True) as r:
            r.raise_for_status()
            with open("stats.details.xlsx", 'wb') as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
        print("\n下载完成!")


import time
start = time.time()
K = ActorTestRange("stable","F:/xiongzhicheng_QM1XZC-O-XiongZhiCheng_2830/")
K.all_to_excel()
end = time.time()
print(end-start)