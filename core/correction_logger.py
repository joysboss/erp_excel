"""
手动指正和日志收集模块
"""
import json
import hashlib
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Any, Optional
import logging

logger = logging.getLogger(__name__)


class CorrectionLogger:
    """手动指正日志记录器"""
    
    def __init__(self, log_dir: str = "data/logs"):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.current_log_file = self.log_dir / f"corrections_{datetime.now().strftime('%Y%m%d')}.jsonl"
    
    def _generate_id(self) -> str:
        """生成日志ID"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        return f"LOG_{timestamp}_{hash(timestamp) % 10000:04d}"
    
    def _get_file_hash(self, headers: List[str]) -> str:
        """生成表头哈希值"""
        headers_str = '|'.join(headers)
        return hashlib.md5(headers_str.encode()).hexdigest()[:16]
    
    def log_correction(
        self,
        headers: List[str],
        auto_mapping: Dict[str, Any],
        manual_correction: Dict[str, Any],
        data_samples: List[Dict[str, Any]] = None,
        reason: str = ""
    ) -> Dict[str, Any]:
        """
        记录手动指正日志
        
        Args:
            headers: 表头列表
            auto_mapping: 自动识别的映射
            manual_correction: 手动指正的映射
            data_samples: 数据样本（前3条）
            reason: 指正原因
        
        Returns:
            日志条目
        """
        log_entry = {
            "id": self._generate_id(),
            "timestamp": datetime.now().isoformat(),
            "file_hash": self._get_file_hash(headers),
            "headers": headers,
            
            "auto_mapping": auto_mapping,
            "manual_correction": manual_correction,
            
            "data_samples": data_samples or [],
            "reason": reason,
            
            # 统计信息
            "correction_count": len(manual_correction),
            "affected_fields": list(manual_correction.keys())
        }
        
        # 追加写入日志文件
        with open(self.current_log_file, 'a', encoding='utf-8') as f:
            f.write(json.dumps(log_entry, ensure_ascii=False) + '\n')
        
        logger.info(f"记录手动指正日志: {log_entry['id']}")
        return log_entry
    
    def load_logs(self, days: int = 30) -> List[Dict[str, Any]]:
        """
        加载最近N天的日志
        
        Args:
            days: 天数
        
        Returns:
            日志列表
        """
        logs = []
        
        for i in range(days):
            log_file = self.log_dir / f"corrections_{datetime.now().strftime('%Y%m%d')}.jsonl"
            if log_file.exists():
                with open(log_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        if line.strip():
                            logs.append(json.loads(line))
        
        return logs
    
    def analyze_logs(self, days: int = 30) -> Dict[str, Any]:
        """
        分析日志，生成优化建议
        
        Args:
            days: 分析最近N天的日志
        
        Returns:
            分析结果
        """
        logs = self.load_logs(days)
        
        if not logs:
            return {
                "total_logs": 0,
                "corrections_by_field": {},
                "new_keywords": [],
                "suggestions": []
            }
        
        # 统计字段指正次数
        corrections_by_field = {}
        new_keywords = {}
        
        for log in logs:
            for field_name in log.get('affected_fields', []):
                if field_name not in corrections_by_field:
                    corrections_by_field[field_name] = {
                        "count": 0,
                        "patterns": {}
                    }
                
                corrections_by_field[field_name]["count"] += 1
                
                # 记录从哪个列名改为哪个列名
                if field_name in log["manual_correction"]:
                    from_col = log["manual_correction"][field_name].get("from", {}).get("column_name", "")
                    to_col = log["manual_correction"][field_name].get("to", {}).get("column_name", "")
                    
                    pattern = f"{from_col} -> {to_col}"
                    if pattern not in corrections_by_field[field_name]["patterns"]:
                        corrections_by_field[field_name]["patterns"][pattern] = 0
                    corrections_by_field[field_name]["patterns"][pattern] += 1
                    
                    # 检测新关键词
                    if to_col not in new_keywords:
                        new_keywords[to_col] = {
                            "field": field_name,
                            "count": 0,
                            "contexts": []
                        }
                    new_keywords[to_col]["count"] += 1
        
        # 生成建议
        suggestions = []
        
        for field_name, field_data in corrections_by_field.items():
            if field_data["count"] >= 3:  # 至少出现3次才建议
                # 找出最常见的模式
                top_patterns = sorted(
                    field_data["patterns"].items(),
                    key=lambda x: x[1],
                    reverse=True
                )[:3]
                
                for pattern, count in top_patterns:
                    from_col, to_col = pattern.split(" -> ")
                    suggestions.append({
                        "field": field_name,
                        "from_column": from_col,
                        "to_column": to_col,
                        "count": count,
                        "action": f"添加'{to_col}'作为{field_name}的关键词",
                        "priority": "high" if count >= 5 else "medium"
                    })
        
        return {
            "total_logs": len(logs),
            "corrections_by_field": corrections_by_field,
            "new_keywords": new_keywords,
            "suggestions": suggestions,
            "analysis_date": datetime.now().isoformat()
        }
    
    def get_stats(self) -> Dict[str, Any]:
        """获取日志统计信息"""
        logs = self.load_logs(30)
        
        total_corrections = sum(log.get("correction_count", 0) for log in logs)
        affected_fields = set()
        
        for log in logs:
            affected_fields.update(log.get("affected_fields", []))
        
        return {
            "total_logs": len(logs),
            "total_corrections": total_corrections,
            "affected_fields": len(affected_fields),
            "fields_list": list(affected_fields),
            "log_dir": str(self.log_dir),
            "current_log_file": str(self.current_log_file)
        }
