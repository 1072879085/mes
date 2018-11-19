package cn.ssm.service.impl;

import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import cn.ssm.mapper.ProductionPlanMapper;
import cn.ssm.mapper.ShopDeliveryMapper;
import cn.ssm.mapper.ShopSortMapper;
import cn.ssm.mapper.ShopTransitionMapper;
import cn.ssm.po.Person;
import cn.ssm.po.ProductionPlan;
import cn.ssm.po.ShopDelivery;
import cn.ssm.po.ShopTransition;
import cn.ssm.util.DataSource;
import cn.ssm.vo.WeekPlan;
import cn.ssm.service.ProductionPlanService;

@Service
@DataSource("dataSource1")
public class ProductionPlanServiceImpl implements ProductionPlanService {
	@Autowired
	private ProductionPlanMapper productionPlanMapper;
	@Autowired
	private ShopDeliveryMapper shopDeliveryMapper;
	@Autowired
	private ShopSortMapper shopSortMapper;
	@Autowired
	private ShopTransitionMapper shopTransitionMapper;
	
	
		
	// 未完成计划单号查询
	@Override
	public String selectJihuadanhao(String clientMaterialNo, String materialNo) {
		List<ProductionPlan> list = productionPlanMapper.selectJihuadanhao(
				clientMaterialNo, materialNo);
		String html = "";
		for (int i = 0; i < list.size(); i++) {
			html += "<option value=" + list.get(i).getPlanNo() + ">"
					+ list.get(i).getPlanNo() + "</option>";
		}
		return html;

	}

	// 计划进度查询
	@Override
	public List<WeekPlan> selectProductionPlanByParam1(String plan_no,
			String client_material_no, String material_no) {
		return productionPlanMapper.selectProductionPlanByParam1(plan_no,
				client_material_no, material_no);
	}

	// 计划进度查询
	@Override
	public Integer selectShopDeliveryByParam(Integer planId) {
		return shopDeliveryMapper.selectShopDeliveryByParam(planId);
	}

	// 计划进度查询
	@Override
	public List<ShopTransition> selectShopTransitionByParam(String plan_no,
			String client_material_no, String material_no) {
		return shopTransitionMapper.selectShopTransitionByParam(plan_no,
				client_material_no, material_no);
	}

	// 计划进度弹窗
	@Override
	public List<ShopTransition> selectShopTransitionByParam1(String plan_no,
			String client_material_no, String material_no) {
		return shopTransitionMapper.selectShopTransitionByParam1(plan_no,
				client_material_no, material_no);
	}

	@Override
	public void addProductionPlan(ProductionPlan productionPlan, Integer num,
			HttpServletRequest request, HttpSession session) {
		List<ShopDelivery> listShopDelivery = new ArrayList<ShopDelivery>();
		String Sort = "";
		if (num != null) {
			for (int i = 1; i <= num; i++) {
				String chejian = request.getParameter("chejian" + i);
				String jiaofu_num = request.getParameter("jiaofu_num" + i);
				String jiaofu_date = request.getParameter("jiaofu_date" + i);
				String plan_finish_date = request
						.getParameter("plan_finish_date" + i);
				String shijiao_num = request.getParameter("shijiao_num" + i);
				String shijiao_date = request.getParameter("shijiao_date" + i);
				ShopDelivery shopDelivery = new ShopDelivery();
				shopDelivery.setShopName(chejian);
				shopDelivery.setDeliveryNum(jiaofu_num);
				shopDelivery.setDeliveryDate(jiaofu_date);
				shopDelivery.setPlanFinishDate(plan_finish_date);
				shopDelivery.setSendNum(shijiao_num);
				shopDelivery.setSendDate(shijiao_date);
				listShopDelivery.add(shopDelivery);
				Sort += chejian + ",";
			}
			Person person = (Person) session.getAttribute("user");
			Sort = Sort.substring(0, Sort.length() - 1);
			if (person != null) {
				productionPlan.setMakePeople(person.getPersonName());
			}
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			String planDate = sf.format(new Date());
			productionPlan.setMakeTime(planDate);
			productionPlan.setSort(Sort);
			productionPlan.setIsMonthlyPlan(0);
			// 表示未完成生产isProduct=0
			productionPlan.setIsProduct(0);
			productionPlan.setIsLatest(1);
			if (productionPlan.getPlanId() == null) {

				productionPlan.setIsNew(1);
				productionPlanMapper.insertSelective(productionPlan);
			} else {
				productionPlan.setIsNew(1);

				productionPlanMapper.updateByPrimaryKey(productionPlan);

			}
			Integer planId = productionPlan.getPlanId();
			shopDeliveryMapper.deleteByKey(planId);
			int a = 0, b = 0, c = 0, d = 0, e= 0;
			for (int i = 0; i < listShopDelivery.size(); i++) {
				ShopDelivery shopDelivery = new ShopDelivery();

				// 判断重复工段，在Priority上+1
				if (listShopDelivery.get(i).getShopName().equals("冲压工段") == true) {
					a += 1;
					listShopDelivery.get(i).setPriority(a);
				} else if (listShopDelivery.get(i).getShopName().equals("仪表工段") == true) {
					b += 1;
					listShopDelivery.get(i).setPriority(b);
				} else if (listShopDelivery.get(i).getShopName().equals("焊接工段") == true) {
					c += 1;
					listShopDelivery.get(i).setPriority(c);
				} else if (listShopDelivery.get(i).getShopName().equals("后道工段") == true) {
					d += 1;
					listShopDelivery.get(i).setPriority(d);
				} else if (listShopDelivery.get(i).getShopName().equals("外协") == true) {
					e += 1;
					listShopDelivery.get(i).setPriority(e);
				} else {
					listShopDelivery.get(i).setPriority(1);
				}

				shopDelivery = listShopDelivery.get(i);
				shopDelivery.setPlanId(planId);
				shopDeliveryMapper.insertSelective(shopDelivery);
			}

		}

	}

	// 周计划记录--未完成
	@Override
	public List<ProductionPlan> selectProductionPlanByParam(int startPos,
			int pageSize, String client, String clientMaterialNo,
			String start_date, String end_date) {

		List<ProductionPlan> listProductionPlan = productionPlanMapper
				.selectProductionPlanByParam(startPos, pageSize, client,
						clientMaterialNo, start_date, end_date);
		return listProductionPlan;
	}

	// 周计划记录--已经完成
	@Override
	public List<ProductionPlan> selectFinishProductionPlanByParam(int startPos,
			int pageSize, String client, String clientMaterialNo,
			String start_date, String end_date) {

		List<ProductionPlan> listProductionPlan = productionPlanMapper
				.selectFinishProductionPlanByParam(startPos, pageSize, client,
						clientMaterialNo, start_date, end_date);
		return listProductionPlan;
	}

	// 月计划记录--未完成
	@Override
	public List<ProductionPlan> selectTotalPlanByParam(int startPos,
			int pageSize, String client, String clientMaterialNo,
			String start_date, String end_date) {

		List<ProductionPlan> listProductionPlan = productionPlanMapper
				.selectTotalPlanByParam(startPos, pageSize, client,
						clientMaterialNo, start_date, end_date);
		return listProductionPlan;
	}

	// 月计划记录--已经完成
	@Override
	public List<ProductionPlan> selectFinishTotalPlanByParam(int startPos,
			int pageSize, String client, String clientMaterialNo,
			String start_date, String end_date) {

		List<ProductionPlan> listProductionPlan = productionPlanMapper
				.selectFinishTotalPlanByParam(startPos, pageSize, client,
						clientMaterialNo, start_date, end_date);
		return listProductionPlan;
	}

	// 月计划or周计划审批
	@Override
	public int updateByKey(Integer planId) {

		return productionPlanMapper.updateByKey(planId);

	}

	@Override
	public ProductionPlan selectProductionPlanByPrimaryKey(Integer planId) {
		if (planId != null) {
			ProductionPlan productionPlan = productionPlanMapper
					.selectByPrimaryKey(planId);
			return productionPlan;
		} else {
			return null;
		}

	}

	@Override
	public List<ShopDelivery> selectShopDeliveryByKey(Integer planId) {
		List<ShopDelivery> listShopDelivery = new ArrayList<ShopDelivery>();
		if (planId != null) {
			listShopDelivery = shopDeliveryMapper.selectByKey(planId);

		} else {
			ShopDelivery shopDelivery = new ShopDelivery();
			shopDelivery.setShopName("");
			listShopDelivery.add(shopDelivery);
		}
		return listShopDelivery;
	}

	@Override
	public void addTotalPlan(ProductionPlan productionPlan, HttpSession session) {
		SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
		String planDate = sf.format(new Date());
		productionPlan.setMakeTime(planDate);
		productionPlan.setIsMonthlyPlan(1);
		productionPlan.setIsNew(1);
		productionPlan.setIsLatest(1);
		// 表示未完成生产isProduct=0
		productionPlan.setIsProduct(0);
		Person person = (Person) session.getAttribute("user");
		if (person != null) {
			productionPlan.setMakePeople(person.getPersonName());
		}
		if (productionPlan.getPlanId() == null) {
			productionPlanMapper.insertSelective(productionPlan);
		} else {
			productionPlanMapper.updateByPrimaryKey(productionPlan);
		}
	}

	@Override
	public List<ShopDelivery> selectShopDeliveryByParam(String materialNo,
			String clientMaterialNo) {
		List<ShopDelivery> listShopDelivery = new ArrayList<ShopDelivery>();
		if (materialNo != null && materialNo != "" && clientMaterialNo != null
				&& clientMaterialNo != "") {
			List<String> shop = shopSortMapper.selectShop(materialNo,
					clientMaterialNo);
			if (shop.size() > 0) {
				for (int i = 0; i < shop.size(); i++) {
					ShopDelivery shopDelivery = new ShopDelivery();
					shopDelivery.setShopName(shop.get(i));
					listShopDelivery.add(shopDelivery);
				}
			} else {
				ShopDelivery shopDelivery = new ShopDelivery();
				shopDelivery.setShopName("");
				listShopDelivery.add(shopDelivery);
			}
		} else {
			ShopDelivery shopDelivery = new ShopDelivery();
			shopDelivery.setShopName("");
			listShopDelivery.add(shopDelivery);

		}
		return listShopDelivery;

	}
	
	
	// 客户查询显示
		@Override
		public String selectClient() {
			List<ProductionPlan> list= productionPlanMapper.selectClient();
			  String html="<select id='client' name='client' class='mini-input multiselect  multiselect2' multiple='multiple'>"+"<option value='"+list.get(0).getClient()+"'  selected='selected' >"+list.get(0).getClient()+"</option>"+"";
			
				for(int i=1;i<list.size();i++){
					html+="<option value="+list.get(i).getClient()+" >"+list.get(i).getClient()+"</option>";												
				}			
			    	html=html+"</select>";
					return html;
		}


}
