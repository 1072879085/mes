package cn.ssm.service;

import java.util.List;

import cn.ssm.po.GetSecDetail;
import cn.ssm.po.GetSecMaterials;
import cn.ssm.po.SecondaryMaterials;
import cn.ssm.vo.InputSec;

public interface SecMaterialService {
	
	 //1.辅料查询分页
	    List<SecondaryMaterials> selectSecByParam(int startPos,int pageSize,String secMaterialsName,String type);
	
	 //2.插入辅料信息
		int insert(SecondaryMaterials record);
		
	 //3.删除辅料信息
		int deleteByPrimaryKey(Integer materialsId);
		
	 //4.更新辅料信息		
		int updateByPrimaryKeySelective(SecondaryMaterials record);
		
		SecondaryMaterials selectByPrimaryKey(Integer materialsId);
		
	 //5.Ajax-excel导入
	    List<SecondaryMaterials> selectSecByParamAjax(String secMaterialsName,String type);
	    
	 //6.领辅料审批记录分页查询
	    List<GetSecMaterials> selectGetSecMaterialsByParam(int startPos,int pageSize,String shopName,String start_date,String end_date);
	  
	 //7.1更新操作（审批）
	    int updateByKey(Integer getMaterialsId,String approver);
	    
	 //7.2更新操作（审批不通过）
	    int updatenotByKey(Integer getMaterialsId,String approver);  
	    
	 //8.查看功能（两个表通过内键getMaterialId查询）
	    GetSecMaterials selectGetSecMaterialsKey(Integer getMaterialsId);
		
		List<GetSecDetail> selectByKey(Integer getMaterialsId);
		
	 //9.领辅料查询：根据车间、起始日期、截止日期、（隐藏条件是否领料？is_return=0）查询
	    List<InputSec> selectinputSecGetMaterial(int startPos,int pageSize,String shopName,String start_date,String end_date);

	 //10.退辅料记录分页
	    List<GetSecDetail> selectoutputSecGetMaterial(int startPos,int pageSize,String reshopName,String start_date,String end_date);
	   
	 //11.领取辅料（excel导出当日的信息为T+模板）
	    List<InputSec> selectexcelinputSecGetMaterial();
	    
	 //12.退取辅料（excel导出当日的信息为T+模板）
	    List<GetSecDetail> selectexceloutputSecGetMaterial();

}
