package cn.ssm.mapper;


import java.util.List;

import org.apache.ibatis.annotations.Param;

import cn.ssm.po.ProductAbnormal;


public interface ProductAbnormalMapper {
    int deleteByPrimaryKey(Integer abnormalId);

    int insert(ProductAbnormal record);

    int insertSelective(ProductAbnormal record);

    ProductAbnormal selectByPrimaryKey(Integer abnormalId);

    int updateByPrimaryKeySelective(ProductAbnormal record);

    int updateByPrimaryKey(ProductAbnormal record);
    
//  查询  
  //产品异常单记录分页
    List<ProductAbnormal> selectProductAbnormalByParam(@Param("startPos")int startPos, @Param("pageSize")int pageSize,@Param("client")String client,@Param("product_name")String product_name);
    
  ////产品异常单记录分页查询返回行数totalCount
  	int selectProductAbnormalByParamtotalCount(@Param("client")String client,@Param("product_name")String product_name);
}