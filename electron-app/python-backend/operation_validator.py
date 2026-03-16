"""
操作校验模块
在执行 Pandas 操作之前，校验依赖是否存在
实现友好的澄清机制
"""

from typing import Dict, List, Any, Optional, Tuple
import pandas as pd


class OperationValidator:
    """操作校验器"""
    
    def __init__(self, tables_meta: Dict[str, Any]):
        """
        初始化校验器
        
        Args:
            tables_meta: 表格元信息
                {
                    "表名": {
                        "columns": ["列1", "列2", ...],
                        "sheets": ["Sheet1", "Sheet2", ...],  # 仅 Excel
                        "dtypes": {"列1": "int64", ...},
                        "row_count": 100
                    }
                }
        """
        self.tables_meta = tables_meta
    
    def validate_operation(self, operation: Dict[str, Any]) -> Dict[str, Any]:
        """
        校验操作意图
        
        Args:
            operation: 操作意图对象
                {
                    "action": "merge" | "filter" | "groupby" | "update" | ...,
                    "table": "表名",
                    "left_table": "左表名",  # merge 操作
                    "right_table": "右表名",  # merge 操作
                    "sheet": "Sheet1",
                    "columns": ["列1", "列2"],
                    "join_key": "关联键",
                    "filter_column": "筛选列",
                    ...
                }
        
        Returns:
            {
                "status": "valid" | "need_clarification" | "error",
                "operation": {...},  # 修正后的操作（如果可以自动修正）
                "questions": [...],  # 需要用户回答的问题
                "message": "...",    # 给用户的消息
            }
        """
        action = operation.get("action", "").lower()
        
        # 根据操作类型分发校验
        validators = {
            "merge": self._validate_merge,
            "update": self._validate_update,
            "filter": self._validate_filter,
            "groupby": self._validate_groupby,
            "select": self._validate_select,
            "export": self._validate_export,
        }
        
        validator = validators.get(action, self._validate_generic)
        return validator(operation)
    
    def _validate_merge(self, op: Dict[str, Any]) -> Dict[str, Any]:
        """校验合并/关联操作"""
        questions = []
        warnings = []
        
        left_table = op.get("left_table")
        right_table = op.get("right_table")
        join_key = op.get("join_key")
        sheet = op.get("sheet")
        mapping = op.get("mapping", {})
        
        # 1. 检查左表是否存在
        if left_table and left_table not in self.tables_meta:
            similar = self._find_similar_table(left_table)
            questions.append({
                "type": "table",
                "field": "left_table",
                "message": f"我没有找到表「{left_table}」",
                "options": list(self.tables_meta.keys()),
                "suggestions": similar,
            })
        
        # 2. 检查右表是否存在
        if right_table and right_table not in self.tables_meta:
            similar = self._find_similar_table(right_table)
            questions.append({
                "type": "table",
                "field": "right_table",
                "message": f"我没有找到表「{right_table}」",
                "options": list(self.tables_meta.keys()),
                "suggestions": similar,
            })
        
        # 3. 检查 Sheet 是否存在
        if sheet and right_table in self.tables_meta:
            table_sheets = self.tables_meta[right_table].get("sheets", [])
            if table_sheets and sheet not in table_sheets:
                if len(table_sheets) == 1:
                    # 只有一个 sheet，自动使用
                    op["sheet"] = table_sheets[0]
                    warnings.append(f"已自动选择 Sheet「{table_sheets[0]}」")
                else:
                    questions.append({
                        "type": "sheet",
                        "field": "sheet",
                        "message": f"在「{right_table}」中没有找到 sheet「{sheet}」",
                        "options": table_sheets,
                        "default": table_sheets[0] if table_sheets else None,
                    })
        
        # 4. 检查关联键是否存在
        if join_key:
            missing_in = []
            
            if left_table in self.tables_meta:
                left_cols = self.tables_meta[left_table].get("columns", [])
                if join_key not in left_cols:
                    similar = self._find_similar_column(join_key, left_cols)
                    missing_in.append({
                        "table": left_table,
                        "similar": similar,
                        "available": left_cols,
                    })
            
            if right_table in self.tables_meta:
                right_cols = self.tables_meta[right_table].get("columns", [])
                if join_key not in right_cols:
                    similar = self._find_similar_column(join_key, right_cols)
                    missing_in.append({
                        "table": right_table,
                        "similar": similar,
                        "available": right_cols,
                    })
            
            if missing_in:
                questions.append({
                    "type": "column",
                    "field": "join_key",
                    "message": f"关联键「{join_key}」在部分表中不存在",
                    "details": missing_in,
                })
        else:
            # 没有指定关联键
            questions.append({
                "type": "column",
                "field": "join_key",
                "message": "请指定用于关联两张表的字段（关联键）",
                "hint": "例如：任务ID、订单号 等",
            })
        
        # 5. 检查映射字段是否存在
        for source_col, target_col in mapping.items():
            if right_table in self.tables_meta:
                right_cols = self.tables_meta[right_table].get("columns", [])
                if source_col not in right_cols:
                    questions.append({
                        "type": "column",
                        "field": f"mapping.{source_col}",
                        "message": f"在「{right_table}」中没有找到字段「{source_col}」",
                        "options": right_cols,
                    })
        
        # 返回结果
        if questions:
            return {
                "status": "need_clarification",
                "operation": op,
                "questions": questions,
                "warnings": warnings,
                "message": "需要补充一些信息才能继续执行",
            }
        
        return {
            "status": "valid",
            "operation": op,
            "warnings": warnings,
            "message": "校验通过" + ("，" + "；".join(warnings) if warnings else ""),
        }
    
    def _validate_update(self, op: Dict[str, Any]) -> Dict[str, Any]:
        """校验更新操作（危险操作，需要确认）"""
        questions = []
        
        table = op.get("table")
        update_columns = op.get("columns", [])
        
        # 检查表是否存在
        if table and table not in self.tables_meta:
            questions.append({
                "type": "table",
                "field": "table",
                "message": f"我没有找到表「{table}」",
                "options": list(self.tables_meta.keys()),
            })
        
        # 检查更新列是否存在
        if table in self.tables_meta:
            table_cols = self.tables_meta[table].get("columns", [])
            for col in update_columns:
                if col not in table_cols:
                    questions.append({
                        "type": "column",
                        "field": "columns",
                        "message": f"在表「{table}」中没有找到字段「{col}」",
                        "options": table_cols,
                    })
        
        # 危险操作确认
        if not op.get("confirmed"):
            questions.append({
                "type": "confirm",
                "field": "confirmed",
                "message": "此操作将修改表中的已有数据",
                "options": [
                    {"value": "overwrite", "label": "覆盖已有值"},
                    {"value": "fill_empty", "label": "仅在为空时补充"},
                    {"value": "cancel", "label": "取消操作"},
                ],
            })
        
        if questions:
            return {
                "status": "need_clarification",
                "operation": op,
                "questions": questions,
                "message": "需要确认更新操作",
            }
        
        return {"status": "valid", "operation": op}
    
    def _validate_filter(self, op: Dict[str, Any]) -> Dict[str, Any]:
        """校验筛选操作"""
        questions = []
        
        table = op.get("table")
        filter_column = op.get("filter_column")
        
        if table and table not in self.tables_meta:
            questions.append({
                "type": "table",
                "field": "table",
                "message": f"我没有找到表「{table}」",
                "options": list(self.tables_meta.keys()),
            })
        
        if filter_column and table in self.tables_meta:
            table_cols = self.tables_meta[table].get("columns", [])
            if filter_column not in table_cols:
                questions.append({
                    "type": "column",
                    "field": "filter_column",
                    "message": f"在表「{table}」中没有找到字段「{filter_column}」",
                    "options": table_cols,
                })
        
        if questions:
            return {
                "status": "need_clarification",
                "operation": op,
                "questions": questions,
            }
        
        return {"status": "valid", "operation": op}
    
    def _validate_groupby(self, op: Dict[str, Any]) -> Dict[str, Any]:
        """校验分组统计操作"""
        questions = []
        
        table = op.get("table")
        group_columns = op.get("group_columns", [])
        agg_columns = op.get("agg_columns", [])
        
        if table and table not in self.tables_meta:
            questions.append({
                "type": "table",
                "field": "table",
                "message": f"我没有找到表「{table}」",
                "options": list(self.tables_meta.keys()),
            })
        
        if table in self.tables_meta:
            table_cols = self.tables_meta[table].get("columns", [])
            
            for col in group_columns + agg_columns:
                if col not in table_cols:
                    questions.append({
                        "type": "column",
                        "field": "columns",
                        "message": f"在表「{table}」中没有找到字段「{col}」",
                        "options": table_cols,
                    })
        
        if questions:
            return {
                "status": "need_clarification",
                "operation": op,
                "questions": questions,
            }
        
        return {"status": "valid", "operation": op}
    
    def _validate_select(self, op: Dict[str, Any]) -> Dict[str, Any]:
        """校验选择列操作"""
        return self._validate_generic(op)
    
    def _validate_export(self, op: Dict[str, Any]) -> Dict[str, Any]:
        """校验导出操作"""
        return {"status": "valid", "operation": op}
    
    def _validate_generic(self, op: Dict[str, Any]) -> Dict[str, Any]:
        """通用校验"""
        questions = []
        
        # 检查涉及的表
        for field in ["table", "left_table", "right_table"]:
            table = op.get(field)
            if table and table not in self.tables_meta:
                questions.append({
                    "type": "table",
                    "field": field,
                    "message": f"我没有找到表「{table}」",
                    "options": list(self.tables_meta.keys()),
                })
        
        if questions:
            return {
                "status": "need_clarification",
                "operation": op,
                "questions": questions,
            }
        
        return {"status": "valid", "operation": op}
    
    def _find_similar_table(self, name: str) -> List[str]:
        """查找相似的表名"""
        similar = []
        name_lower = name.lower()
        
        for table in self.tables_meta.keys():
            table_lower = table.lower()
            # 简单的相似度判断
            if name_lower in table_lower or table_lower in name_lower:
                similar.append(table)
            elif self._levenshtein_distance(name_lower, table_lower) <= 3:
                similar.append(table)
        
        return similar[:3]  # 最多返回 3 个建议
    
    def _find_similar_column(self, name: str, columns: List[str]) -> List[str]:
        """查找相似的列名"""
        similar = []
        name_lower = name.lower()
        
        for col in columns:
            col_lower = col.lower()
            if name_lower in col_lower or col_lower in name_lower:
                similar.append(col)
            elif self._levenshtein_distance(name_lower, col_lower) <= 2:
                similar.append(col)
        
        return similar[:3]
    
    @staticmethod
    def _levenshtein_distance(s1: str, s2: str) -> int:
        """计算编辑距离"""
        if len(s1) < len(s2):
            return OperationValidator._levenshtein_distance(s2, s1)
        
        if len(s2) == 0:
            return len(s1)
        
        previous_row = range(len(s2) + 1)
        
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        
        return previous_row[-1]


def format_clarification_message(result: Dict[str, Any]) -> str:
    """
    将校验结果格式化为用户友好的消息
    """
    if result["status"] == "valid":
        return result.get("message", "✅ 校验通过")
    
    lines = ["⚠️ 需要补充一些信息：", ""]
    
    for i, q in enumerate(result.get("questions", []), 1):
        msg = q.get("message", "")
        lines.append(f"{i}. {msg}")
        
        options = q.get("options", [])
        if options:
            if len(options) <= 5:
                lines.append(f"   可选：{', '.join(str(o) for o in options)}")
            else:
                lines.append(f"   可选（共 {len(options)} 项）：{', '.join(str(o) for o in options[:5])}...")
        
        suggestions = q.get("suggestions", [])
        if suggestions:
            lines.append(f"   建议：{', '.join(suggestions)}")
        
        lines.append("")
    
    return "\n".join(lines)


# 意图识别失败时的友好消息
CLARIFICATION_MESSAGES = {
    "no_intent": {
        "message": "我没有识别到明确的表处理操作",
        "suggestions": [
            "筛选【表名】中【条件】的记录",
            "统计【表名】中【字段】的数量",
            "从【表A】补充【字段】到【表B】",
        ],
        "hint": "你是想【筛选 / 统计 / 合并 / 导出】哪一种？",
    },
    "ambiguous_table": {
        "message": "同时命中了多个表，请明确指定",
    },
    "out_of_scope": {
        "message": "当前模式为【表处理模式】",
        "hint": "你的问题更像是业务咨询或分析说明，暂不支持直接执行",
        "suggestions": [
            "切换为【查询说明】模式",
            "用一句话明确你希望对表做的操作",
        ],
    },
}


def get_clarification_response(error_type: str, **kwargs) -> Dict[str, Any]:
    """
    获取标准化的澄清响应
    """
    template = CLARIFICATION_MESSAGES.get(error_type, {})
    
    return {
        "status": "need_clarification",
        "type": error_type,
        "message": template.get("message", "需要更多信息"),
        "suggestions": template.get("suggestions", []),
        "hint": template.get("hint", ""),
        "options": kwargs.get("options", []),
        **kwargs,
    }































