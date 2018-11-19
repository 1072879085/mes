package cn.ssm.service.impl;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import cn.ssm.mapper.DepOpinionMapper;
import cn.ssm.mapper.ProductAbnormalMapper;
import cn.ssm.mapper.ProductionPlanMapper;
import cn.ssm.po.DepOpinion;
import cn.ssm.po.ProductAbnormal;
import cn.ssm.po.ProductionPlan;
import cn.ssm.service.ProductAbnormalService;
import cn.ssm.util.DataSource;

@Service
@DataSource("dataSource1")
public class ProductAbnormalServiceImpl implements ProductAbnormalService {
	@Autowired
	private ProductAbnormalMapper productAbnormalMapper;
	@Autowired
	private DepOpinionMapper  depOpinionMapper;
	@Autowired
	private ProductionPlanMapper productionPlanMapper;
	
	
 //	查询
	//产品异常单记录分页
	@Override
	public List<ProductAbnormal> selectProductAbnormalByParam(int startPos,
			int pageSize, String client, String product_name) {
		return productAbnormalMapper.selectProductAbnormalByParam(startPos, pageSize,client,product_name);
	}
	
	
	
	
	
//添加
	@Override
	public void insertProductAbnormal(ProductAbnormal productAbnormal){
		 String clientMaterialNo=productAbnormal.getClientMaterialNo();
		 String materialNo=productAbnormal.getMaterialNo();
		//从ProductionPlan中的取出计划单号，产品名称，产品规格
			ProductionPlan productionPlan= new ProductionPlan();
			productionPlan=productionPlanMapper.selectJCC(clientMaterialNo, materialNo);
			//再将计划单号set到GetMaterial
		if(clientMaterialNo!=null&&!("".equals(clientMaterialNo))&&materialNo!=null&&!("".equals(materialNo))){
				if(productionPlan!=null&&!("".equals(productionPlan))){
					productAbnormal.setProductName(productionPlan.getProductName());
					productAbnormal.setProductSpecfic(productionPlan.getProductSpec());
				}else{
					productAbnormal.setProductName("");
					productAbnormal.setProductSpecfic("");
				}
			
		}else{
			productAbnormal.setProductName("");
			productAbnormal.setProductSpecfic("");
		}	
		 productAbnormalMapper.insertSelective(productAbnormal);
	}
	@Override
	public void insertAbnormalId(List<DepOpinion> listDepOpinion, Integer abnormalId){
		for(int i=0;i<listDepOpinion.size();i++){
			   DepOpinion depOpinion=new DepOpinion();
			   depOpinion=listDepOpinion.get(i);
			   depOpinion.setAbnormalId(abnormalId);
			   depOpinionMapper.insertSelective(depOpinion);
		   }
	}
	
	
	
	
	
//修改
	@Override
	public ProductAbnormal selectByPrimaryKey(Integer abnormalId){	
		return productAbnormalMapper.selectByPrimaryKey(abnormalId);
	}
	@Override
	public List<DepOpinion> selectByAbnormalId(Integer abnormalId){
		return depOpinionMapper.selectByAbnormalId(abnormalId);
	}
	@Override
	public void deleteByAbnormalId(Integer abnormalId){
		depOpinionMapper.deleteByAbnormalId(abnormalId);
	}
	@Override
	public int updateByPrimaryKeySelective(ProductAbnormal productAbnormal){	 
		return productAbnormalMapper.updateByPrimaryKeySelective(productAbnormal);
	 }
	@Override
	public void updateAbnormalId(List<DepOpinion> listDepOpinion, Integer abnormalId){
		 for(int i=0;i<listDepOpinion.size();i++){
			   DepOpinion depOpinion=new DepOpinion();
			   depOpinion=listDepOpinion.get(i);
			   depOpinion.setAbnormalId(abnormalId);
			   depOpinionMapper.insertSelective(depOpinion);
		   }
		
	}


  
    
}
