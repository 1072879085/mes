package cn.ssm.controller;

import java.io.File;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import javax.swing.filechooser.FileSystemView;

import jxl.Cell;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import cn.ssm.mapper.DailyCheckMapper;
import cn.ssm.mapper.GetMaterialMapper;
import cn.ssm.mapper.GetSecDetailMapper;
import cn.ssm.mapper.GetSecMaterialsMapper;
import cn.ssm.mapper.MoldRecordMapper;
import cn.ssm.mapper.ProductAbnormalMapper;
import cn.ssm.mapper.ProductRecordMapper;
import cn.ssm.mapper.ProductTestMapper;
import cn.ssm.mapper.ProductionPlanMapper;
import cn.ssm.mapper.ShopDeliveryMapper;
import cn.ssm.mapper.ShopPlanMapper;
import cn.ssm.mapper.ShopTransitionMapper;
import cn.ssm.mapper.SpcMapper;
import cn.ssm.mapper.TrackCardMapper;
import cn.ssm.mapper.WorkCardMapper;
import cn.ssm.po.CheckRecord;
import cn.ssm.po.Cipin;
import cn.ssm.po.DailyCheck;
import cn.ssm.po.DepOpinion;
import cn.ssm.po.GetDetail;
import cn.ssm.po.GetMaterial;
import cn.ssm.po.GetSecDetail;
import cn.ssm.po.GetSecMaterials;
import cn.ssm.po.MoldRecord;
import cn.ssm.po.Person;
import cn.ssm.po.ProductAbnormal;
import cn.ssm.po.ProductRecord;
import cn.ssm.po.ProductTest;
import cn.ssm.po.ProductionPlan;
import cn.ssm.po.ShopPlan;
import cn.ssm.po.ShopTransition;
import cn.ssm.po.Spc;
import cn.ssm.po.SpcTest;
import cn.ssm.po.TrackCard;
import cn.ssm.po.WorkCard;
import cn.ssm.service.DailyCheckService;
import cn.ssm.service.GetDetailService;
import cn.ssm.service.GetMaterialService;
import cn.ssm.service.IntputOrOutputRecordService;
import cn.ssm.service.MoldRecordService;
import cn.ssm.service.ProductAbnormalService;
import cn.ssm.service.ProductTestService;
import cn.ssm.service.ProductionPlanService;
import cn.ssm.service.SalaryService;
import cn.ssm.service.SecMaterialService;
import cn.ssm.service.ShopPlanService;
import cn.ssm.service.SpcService;
import cn.ssm.service.TJdbcService;
import cn.ssm.service.TrackCardSelectService;
import cn.ssm.service.WorkCardService;
import cn.ssm.util.Page;
import cn.ssm.vo.ExterAssociation;
import cn.ssm.vo.Input;
import cn.ssm.vo.InputMaterialAssociation;
import cn.ssm.vo.InputSec;
import cn.ssm.vo.Output;
import cn.ssm.vo.ProductionRecordInquiry;
import cn.ssm.vo.Salary;
import cn.ssm.vo.TemPrice;
import cn.ssm.vo.WeekPlan;

@Controller
@RequestMapping("/record")
public class RecordController {
	@Autowired
	private ProductionPlanService productionPlanService;	
	@Autowired
	private ShopPlanService shopPlanService;	
	@Autowired
	private WorkCardService workCardService;
	@Autowired
	private IntputOrOutputRecordService inputOrOutputRecordService;
	@Autowired
	private SalaryService salaryService;
	@Autowired
	private GetMaterialService getMaterialService;
	@Autowired
	private MoldRecordService moldRecordService;
	@Autowired
	private ProductTestService productTestService;
	@Autowired
	private TrackCardSelectService trackCardSelectService;
	@Autowired
	private ProductAbnormalService productAbnormalService;
	@Autowired
	private DailyCheckService dailyCheckService;
	@Autowired
	private GetDetailService getDetailService;
	@Autowired
	private SpcService spcService;
	@Autowired
	private SecMaterialService secMaterialService;
	@Autowired
	private TJdbcService tJdbcService;

	// 分页所需的
	@Autowired
	private ProductionPlanMapper productionPlanMapper;
	@Autowired
	private GetMaterialMapper getMaterialMapper;
	@Autowired
	private ShopPlanMapper shopPlanMapper;
	@Autowired
	private ProductAbnormalMapper productAbnormalMapper;
	@Autowired
	private TrackCardMapper trackCardMapper;
	@Autowired
	private WorkCardMapper workCardMapper;
	@Autowired
	private MoldRecordMapper moldRecordMapper;
	@Autowired
	private ProductTestMapper productTestMapper;
	@Autowired
	private DailyCheckMapper dailycheckMapper;
	@Autowired
	private ProductRecordMapper productRecordMapper;
	@Autowired
	private SpcMapper spcMapper;
	@Autowired
	private GetSecMaterialsMapper getSecMaterialsMapper;
	@Autowired
	private GetSecDetailMapper getSecDetailMapper;
	@Autowired
	private ShopTransitionMapper shopTransitionMapper;
	@Autowired
	private ShopDeliveryMapper shopDeliveryMapper ;
	
	// 首页跳转
	@RequestMapping("/toHomePage")
	public String toHomePage() throws Exception {

		return "homePage";
	}

	// 计划进度查询跳转页面
	@RequestMapping("/toselectProductionPlan")
	public String toselectProductionPlan() throws Exception {
		return "productionproess";
	}

	// 半成品入库记录问题是否解决（将已解决由0到1）
	@RequestMapping("/approveMiddleProblem")
	public String approveProblem(ProductRecord productRecord, Integer recordId,
			String temPrice, HttpSession session, Model model) throws Exception {

		inputOrOutputRecordService.updateIsProblem(recordId);
		return "redirect:OutputMiddleMaterialsRecord.action";

	}

	// 成品入库记录问题是否解决（将已解决由0到1）
	@RequestMapping("/approveFullProblem")
	public String approveFullProblem(ProductRecord productRecord,
			Integer recordId, String temPrice, HttpSession session, Model model)
			throws Exception {

		inputOrOutputRecordService.updateIsProblem(recordId);
		return "redirect:toOutputFullMaterialsRecord.action";

	}

	// 成品入库===========================================
	@RequestMapping("/toOutputFullMaterialsRecord")
	public String toOutputFullMaterialsRecord(Integer pageNow,
			String material_no, String start_date, String end_date, Model model) {
		// 第一、二张表的查询
		List<ProductRecord> listoutputmaterial = new ArrayList<ProductRecord>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = productRecordMapper.selectOutputFullRecordtotalCount(
				material_no, start_date, end_date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listoutputmaterial = inputOrOutputRecordService.selectOutputFullRecord(
				page.getStartPos(), page.getPageSize(), material_no,
				start_date, end_date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listoutputmaterial", listoutputmaterial);
		model.addAttribute("page", page);
		model.addAttribute("material_no", material_no);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		return "outputfull";

	}

	// 成品入库===========================================
	@RequestMapping("toOutputFullRecord")
	public String toOutputFullRecord(Integer recordId, Model model)
			throws Exception {
		// 第一、二张表GetMaterial的查询
		ProductRecord outputmiddlerecord = new ProductRecord();
		outputmiddlerecord = inputOrOutputRecordService
				.selectMiddleId(recordId);
		model.addAttribute("outputmiddlerecord", outputmiddlerecord);
		return "lookoutputfull";
	}

	// 1、临时工价审批查询功能：分页查询
	// 临时工价审批记录分页查询
	@RequestMapping("/toTemprice")
	public String toTemprice(HttpSession session, String shopName,
			String processName, String start_date, String end_date,
			Integer pageNow, Model model) throws Exception {
		List<TemPrice> listtemprice = new ArrayList<TemPrice>();
		int totalCount = 0;

		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}
		// 领料审批记录分页查询返回行数totalCount
		totalCount = trackCardMapper.SelectByTempricetotalCount(shopName,
				processName, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listtemprice = workCardService
				.SelectByTemprice(page.getStartPos(), page.getPageSize(),
						shopName, processName, start_date, end_date);

		// controller层中使model能返回一个jsp页面
		model.addAttribute("listtemprice", listtemprice);
		model.addAttribute("shopName", shopName);
		model.addAttribute("processName", processName);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);
		model.addAttribute("page", page);

		// (更新查询数量的变化)车间主任需求：添加对审批数量的查询
		Integer num3 = trackCardMapper.selectTempriceCount();
		session.setAttribute("num3", num3);

		return "approve_temprice";
	}

	// 2审批or不审批功能

	// 2.1审批功能（将已审批由0到1）

	@RequestMapping("/approveTemprice")
	public String approveTemprice(WorkCard workCard, String batchNo,
			String temPrice, HttpSession session, Model model) throws Exception {
		workCardService.updateapproveTemprice(batchNo, temPrice);
		return "redirect:toTemprice.action";
	}

	// 2.2审批功能（将审批不通过由0到2）

	@RequestMapping("/notapproveTemprice")
	public String notapproveTemprice(WorkCard workCard, String batchNo,
			HttpSession session, Model model) throws Exception {

		workCardService.updatenotapproveTemprice(batchNo);
		return "redirect:toTemprice.action";

	}

	// Spc记录(查询)
	@RequestMapping("/toSpcrecord")
	public String toSpcrecord(Integer pageNow, String materialNo,
			String batchNo, String process, String characterVal, Model model) {
		List<Spc> listSpcrecord = new ArrayList<Spc>();
		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = spcMapper.selectSpcBytotalCount(materialNo, batchNo,
				process, characterVal);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listSpcrecord = spcService.selectSpcrecord(page.getStartPos(),
				page.getPageSize(), materialNo, batchNo, process, characterVal);
		model.addAttribute("listSpcrecord", listSpcrecord);
		model.addAttribute("page", page);
		model.addAttribute("materialNo", materialNo);
		model.addAttribute("batchNo", batchNo);
		model.addAttribute("process", process);
		model.addAttribute("characterVal", characterVal);
		return "Spc_record";
	}

	// Spc记录的修改操作
	// 先查询出来，用于数据的回显
	@RequestMapping("/toupdateSpcrecord")
	public String toupdateSpcrecord(HttpServletRequest request, String batchNo,
			String process, String characterVal, Model model) throws Exception {
		// 查询Spc表的主要信息并回显
		Spc spc = spcService.selectUpdateSpcrecord(batchNo, process,
				characterVal);
		model.addAttribute("spc", spc);
		// 查询Spc_test的主要信息并回显
		List<SpcTest> listSpcTest = spcService.selectEditSpcrecord(batchNo,
				process, characterVal);
		model.addAttribute("listSpcTest", listSpcTest);
		return "editSpc_record";
	}

	@RequestMapping("/saveOrupdateSpc")
	public String saveOrupdateSpc(HttpServletRequest request, Spc spcId,
			Model model) throws Exception {
		List<SpcTest> listSpcTest = new ArrayList<SpcTest>();
		for (int i = 1; i <= 25; i++) {
			// 公共部分
			String clientMaterialNo = request.getParameter("clientMaterialNo");
			String materialNo = request.getParameter("materialNo");
			String batchNo = request.getParameter("batchNo");
			String processName = request.getParameter("processName");
			String characterVal = request.getParameter("characterVal");

			// 循环的部分
			String testId = request.getParameter("testId" + i);
			String testVal1 = request.getParameter("testVal1" + i);
			String testVal2 = request.getParameter("testVal2" + i);
			String testVal3 = request.getParameter("testVal3" + i);
			String testVal4 = request.getParameter("testVal4" + i);
			String testVal5 = request.getParameter("testVal5" + i);
			SpcTest s = new SpcTest();

			// 公共部分插入
			s.setMaterialNo(materialNo);
			s.setClientMaterialNo(clientMaterialNo);
			s.setBatchNo(batchNo);
			s.setProcessName(processName);
			s.setCharacterVal(characterVal);

			// 循环的部分插入
			s.setTestId(Integer.valueOf(testId));
			s.setTestVal1(Float.valueOf(testVal1));
			s.setTestVal2(Float.valueOf(testVal2));
			s.setTestVal3(Float.valueOf(testVal3));
			s.setTestVal4(Float.valueOf(testVal4));
			s.setTestVal5(Float.valueOf(testVal5));

			Float TestVal1 = Float.valueOf(testVal1);
			Float TestVal2 = Float.valueOf(testVal2);
			Float TestVal3 = Float.valueOf(testVal3);
			Float TestVal4 = Float.valueOf(testVal4);
			Float TestVal5 = Float.valueOf(testVal5);
			double sumX = (double) Math.round((TestVal1 + TestVal2 + TestVal3
					+ TestVal4 + TestVal5) * 100) / 100;
			s.setSumX(sumX);
			// 插入平均值
			s.setAveX((double) Math.round((sumX / 5) * 100) / 100);
			// 插入方差(取得这5个测量值每组的最大值和最小值)
			float max = 0;
			float min = 0;

			// 找出最大值赋给max
			if (TestVal1 >= TestVal2 && TestVal1 >= TestVal3
					&& TestVal1 >= TestVal4 && TestVal1 >= TestVal5) {
				max = TestVal1;
			} else if (TestVal2 >= TestVal1 && TestVal2 >= TestVal3
					&& TestVal2 >= TestVal4 && TestVal2 >= TestVal5) {
				max = TestVal2;
			} else if (TestVal3 >= TestVal1 && TestVal3 >= TestVal2
					&& TestVal3 >= TestVal4 && TestVal3 >= TestVal5) {
				max = TestVal3;
			} else if (TestVal4 >= TestVal1 && TestVal4 >= TestVal2
					&& TestVal4 >= TestVal3 && TestVal4 >= TestVal5) {
				max = TestVal4;
			} else if (TestVal5 >= TestVal1 && TestVal5 >= TestVal2
					&& TestVal5 >= TestVal3 && TestVal5 >= TestVal4) {
				max = TestVal5;
			} else {
				max = TestVal5;
			}
			// 找出最小值赋给min
			if (TestVal1 <= TestVal2 && TestVal1 <= TestVal3
					&& TestVal1 <= TestVal4 && TestVal1 <= TestVal5) {
				min = TestVal1;
			} else if (TestVal2 <= TestVal1 && TestVal2 <= TestVal3
					&& TestVal2 <= TestVal4 && TestVal2 <= TestVal5) {
				min = TestVal2;
			} else if (TestVal3 <= TestVal1 && TestVal3 <= TestVal2
					&& TestVal3 <= TestVal4 && TestVal3 <= TestVal5) {
				min = TestVal3;
			} else if (TestVal4 <= TestVal1 && TestVal4 <= TestVal2
					&& TestVal4 <= TestVal3 && TestVal4 <= TestVal5) {
				min = TestVal4;
			} else if (TestVal5 <= TestVal1 && TestVal5 <= TestVal2
					&& TestVal5 <= TestVal3 && TestVal5 <= TestVal4) {
				min = TestVal5;
			} else {
				min = TestVal5;
			}
			s.setR(Double.parseDouble(String.valueOf(max))
					- Double.parseDouble(String.valueOf(min)));
			listSpcTest.add(s);
		}
		// 先更新spc表
		spcService.updateByPrimaryKeySelective(spcId);
		// 再更新spcTest
		spcService.updateSpcTest(listSpcTest);
		return "redirect:toSpcrecord.action";
	}

	// 材料批次号显示
	@RequestMapping("/cailiaopicihaoAjax")
	public void cailiaopicihaoAjax(String clientMaterialNo, String materialNo,
			String batchNo, HttpServletRequest request,
			HttpServletResponse response, Model model) throws Exception {
		response.setContentType("text/html;charset=utf-8");
		request.setCharacterEncoding("utf-8");
		String html = getDetailService.selectCailiaopicihao(clientMaterialNo,
				materialNo, batchNo);
		PrintWriter out = response.getWriter();
		out.print(html);

	}
	
	// 客户查询显示
		@RequestMapping("/clientAjax")
		public void clientAjax(HttpServletRequest request,HttpServletResponse response, Model model) throws Exception {
			response.setContentType("text/html;charset=utf-8");
			request.setCharacterEncoding("utf-8");
			String html = productionPlanService.selectClient();
			PrintWriter out = response.getWriter();
			out.print(html);

		}


	// 未完成计划单号查询
	@RequestMapping("/jihuadanhaoAjax")
	public void jihuadanhaoAjax(String clientMaterialNo, String materialNo,
			HttpServletRequest request, HttpServletResponse response,
			Model model) throws Exception {
		response.setContentType("text/html;charset=utf-8");
		request.setCharacterEncoding("utf-8");
		String html = productionPlanService.selectJihuadanhao(clientMaterialNo,
				materialNo);
		PrintWriter out = response.getWriter();
		out.print(html);

	}

	// 故障查询
	@RequestMapping("/tobreakdown")
	public String tobreakdown(Integer pageNow, String batchNo,
			String processName, String assetNo, Model model) throws Exception {
		List<DailyCheck> listDailyCheck = new ArrayList<DailyCheck>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = dailycheckMapper.selectByPrimarytotalCount1(batchNo,
				processName, assetNo);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listDailyCheck = dailyCheckService.selectByPrimary1(page.getStartPos(),
				page.getPageSize(), batchNo, processName, assetNo);
		model.addAttribute("listDailyCheck", listDailyCheck);
		model.addAttribute("page", page);
		model.addAttribute("batchNo", batchNo);
		model.addAttribute("processName", processName);
		model.addAttribute("assetNo", assetNo);
		return "breakdown";
	}

	// 次品记录记录单
	@RequestMapping("/selectWorkCard")
	public String SelectWorkCard(Integer pageNow, String produceDate,
			String produceDate1, Model model) throws Exception {
		List<WorkCard> listWorkCard = new ArrayList<WorkCard>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = workCardMapper.selectWorkCardByParamtotalCount(
				produceDate, produceDate1);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listWorkCard = workCardService.selectWorkCardByParam(
				page.getStartPos(), page.getPageSize(), produceDate,
				produceDate1);
		model.addAttribute("listWorkCard", listWorkCard);
		model.addAttribute("page", page);
		model.addAttribute("produceDate", produceDate);
		model.addAttribute("produceDate1", produceDate1);

		return "workCard";

	}

	// 计划进度查询
	@RequestMapping("/selectProductionPlan")
	public String SelectProductionPlan(String plan_no,
			String client_material_no, String material_no, Model model)
			throws Exception {
		List<ShopTransition> listShopTransition1 = new ArrayList<ShopTransition>();
		listShopTransition1 = productionPlanService
				.selectShopTransitionByParam1(plan_no, client_material_no,
						material_no);
		//System.out.println(listShopTransition1.size());
		

		List<WeekPlan> listWeekPlan = new ArrayList<WeekPlan>();
		listWeekPlan = productionPlanService.selectProductionPlanByParam1(
				plan_no, client_material_no, material_no);
		if(listShopTransition1.size()==0){
			
		String	batchNo=shopPlanService.selectBatchNoByPlanNo(plan_no);
		
		listWeekPlan.get(0).setBatchNo(batchNo);
	
			Integer planId1 = listWeekPlan.get(0).getPlanId();
			
		 String shopName=shopDeliveryMapper.selectShopNameByPlanId(planId1);
		 listWeekPlan.get(0).setShopName(shopName);
				   
			model.addAttribute("listWeekPlan", listWeekPlan);
			
			return "productionproess1";
			
		}else {
			if (listWeekPlan.size() > 0) {
				// 集合中获得固定元素用 .get(0).getX
				Integer planId = listWeekPlan.get(0).getPlanId();
				Integer sendNum = null;
				if (planId != null)
					// 通过PlanId查询sendNum实交数量
					sendNum = productionPlanService
					.selectShopDeliveryByParam(planId);
				String sendNum2 = "";
				if (sendNum != null)
					sendNum2 = sendNum.toString();
				
				List<ShopTransition> listShopTransition = new ArrayList<ShopTransition>();
				listShopTransition = productionPlanService
						.selectShopTransitionByParam(plan_no, client_material_no,
								material_no);
				List<WeekPlan> listWeekPlan2 = new ArrayList<WeekPlan>();
				for (int i = 0; i < listShopTransition.size(); i++) {
					WeekPlan weekPlan = new WeekPlan();
					weekPlan.setShop2(listShopTransition.get(i).getShop2());
					weekPlan.setClient(listWeekPlan.get(0).getClient());
					weekPlan.setClientMaterialNo(listWeekPlan.get(0)
							.getClientMaterialNo());
					weekPlan.setMaterialNo(listWeekPlan.get(0).getMaterialNo());
					weekPlan.setOrderDate(listWeekPlan.get(0).getOrderDate());
					weekPlan.setPlanId(listWeekPlan.get(0).getPlanId());
					weekPlan.setPlanNo(listWeekPlan.get(0).getPlanNo());
					weekPlan.setPlanNum(listWeekPlan.get(0).getPlanNum());
					weekPlan.setQualifiedNum(listShopTransition.get(i)
							.getQualifiedNum());
					weekPlan.setSendNum(sendNum2);
					weekPlan.setBatchNo(listShopTransition.get(i).getBatchNo());
					weekPlan.setAcceptor(listShopTransition.get(i).getAcceptor());
					weekPlan.setShop1(listShopTransition.get(i).getShop1());
					weekPlan.setProvider(listShopTransition.get(i).getProvider());
					weekPlan.setSort(listWeekPlan.get(0).getSort());
					listWeekPlan2.add(weekPlan);
				}
				
				model.addAttribute("listWeekPlan", listWeekPlan2);
				
			}
			model.addAttribute("listShopTransition1", listShopTransition1);
		
			return "productionproess";
		}
		
		
		
	}

	// 产品异常单的查询
	// 产品异常单记录分页
	@RequestMapping("/selectProductAbnormal")
	public String selectProductAbnormal(Integer pageNow, String product_name,
			String client, Model model) throws Exception {

		List<ProductAbnormal> listProductAbnormal = new ArrayList<ProductAbnormal>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = productAbnormalMapper
				.selectProductAbnormalByParamtotalCount(product_name, client);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listProductAbnormal = productAbnormalService
				.selectProductAbnormalByParam(page.getStartPos(),
						page.getPageSize(), product_name, client);

		model.addAttribute("listProductAbnormal", listProductAbnormal);
		model.addAttribute("page", page);
		model.addAttribute("product_name", product_name);
		model.addAttribute("client", client);

		return "ProductAbnormal";
	}

	// 产品异常单的修改

	@RequestMapping("/getProductAbnormal")
	public String getProductAbnormal(Integer abnormalId, Model model)
			throws Exception {
		ProductAbnormal productAbnormal = productAbnormalService
				.selectByPrimaryKey(abnormalId);
		model.addAttribute("productAbnormal", productAbnormal);
		List<DepOpinion> listDepOpinion = new ArrayList<DepOpinion>();
		listDepOpinion = productAbnormalService.selectByAbnormalId(abnormalId);
		model.addAttribute("listDepOpinion", listDepOpinion);
		return "editproductabnormal";
	}

	@RequestMapping("/updateProductAbnormal")
	public String updateProductAbnormal(HttpServletRequest request,
			ProductAbnormal productAbnormal, Integer num1, Model model)
			throws Exception {
		List<DepOpinion> listDepOpinion = new ArrayList<DepOpinion>();

		for (int i = 1; i <= num1; i++) {
			String department = request.getParameter("department" + i);
			String opinion = request.getParameter("opinion" + i);

			DepOpinion depOpinion = new DepOpinion();

			depOpinion.setDepartment(department);
			depOpinion.setOpinion(opinion);

			listDepOpinion.add(depOpinion);

		}

		productAbnormalService.updateByPrimaryKeySelective(productAbnormal);
		Integer abnormalId = productAbnormal.getAbnormalId();
		productAbnormalService.deleteByAbnormalId(abnormalId);
		productAbnormalService.updateAbnormalId(listDepOpinion, abnormalId);
		return "redirect:selectProductAbnormal";

	}

	// 跟踪单四个条件查询
	// 跟踪单记录分页
	@RequestMapping("/getTrackCard")
	public String getTrackCard(Integer pageNow,HttpSession session, String client, String plan_no,
			String client_material_no, String material_no, Model model)
			throws Exception {
		// 定义一个参数listTrackCard作为调用service层方法 的返回值
		List<TrackCard> listTrackCard = new ArrayList<TrackCard>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = trackCardMapper.selectTrackCardByParamtotalCount(client,
				plan_no, client_material_no, material_no);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		// 调用service的层中实现类的方法
		listTrackCard = trackCardSelectService.selectTrackCardByParam(
				page.getStartPos(), page.getPageSize(), client, plan_no,
				client_material_no, material_no);
		// controller层中使model能返回一个jsp页面，在jsp页面便于用EL表达式${}获取值，使model.addAttribute返回jsp页面输出
		model.addAttribute("listTrackCard", listTrackCard);
		model.addAttribute("page", page);
		model.addAttribute("client", client);
		model.addAttribute("plan_no", plan_no);
		model.addAttribute("client_material_no", client_material_no);
		model.addAttribute("material_no", material_no);
		return "genzongdanlist";
	}

	// 跟踪单的修改操作
	// 1.先查询三表进行回显
	@RequestMapping("/UpdateTrackCard")
	public String UpdateTrackCard(Integer cardId, Model model) throws Exception {
		// TrackCard表的查询（1表）
		TrackCard trackCard = trackCardSelectService.selectByPrimaryKey(cardId);
		model.addAttribute("trackCard", trackCard);
		// WorkCard表的查询（2表）
		List<WorkCard> listWorkCard = new ArrayList<WorkCard>();
		listWorkCard = trackCardSelectService.SelectBycardId(cardId);
		model.addAttribute("listWorkCard", listWorkCard);
		return "editgenzongdan";
	}

	// 2.再进行修改（更新）操作
	@RequestMapping("/editTrackCard")
	public String editTrackCard(HttpServletRequest request,
			TrackCard trackCard, Integer num1, Model model) {
		// 循环获得二表的信息
		List<WorkCard> listWorkCard = new ArrayList<WorkCard>();
		if (num1 != null) {
			for (int i = 1; i <= num1; i++) {
				// 获得前端的信息
				String trackId = request.getParameter("trackId" + i);
				String cardId3 = request.getParameter("cardId" + i);
				String shopName = request.getParameter("shopName" + i);
				String processName = request.getParameter("processName" + i);
				String operator = request.getParameter("operator" + i);
				String asset = request.getParameter("asset" + i);
				String assetState = request.getParameter("assetState" + i);
				String mold = request.getParameter("mold" + i);
				String moldState = request.getParameter("moldState" + i);
				String totalNum = request.getParameter("totalNum" + i);
				String hegeNum = request.getParameter("hegeNum" + i);
				String produceDate = request.getParameter("produceDate" + i);
				String price = request.getParameter("price" + i);
				// 新建类接受前端的信息，把信息放到类中
				WorkCard workCard = new WorkCard();

				// trackId加上强制转换（从String到Int型）
				Integer trackId1 = Integer.parseInt(trackId);
				workCard.setTrackId(trackId1);
				Integer cardId1 = Integer.parseInt(cardId3);
				workCard.setCardId(cardId1);
				workCard.setShopName(shopName);
				workCard.setProcessName(processName);
				workCard.setOperator(operator);
				workCard.setAsset(asset);
				workCard.setAssetState(assetState);
				workCard.setMold(mold);
				workCard.setMoldState(moldState);
				workCard.setTotalNum(totalNum);
				workCard.setHegeNum(hegeNum);
				workCard.setProduceDate(produceDate);
				workCard.setPrice(price);
				// 把类放到集合中去
				listWorkCard.add(workCard);
				String temp2 = "num" + (i + 1);
				// 获得前端num2,3,4....的值
				String temp3 = request.getParameter(temp2);
				// 判断num是否为空，防止为空时为空字符串，无次品时不执行下列循环
				if (temp3 != null && temp3 != "") {
					// String强制转换成int
					Integer num2 = Integer.parseInt(temp3);
					List<Cipin> listCipin = new ArrayList<Cipin>();
					for (int j = 1; j <= num2; j++) {
						// i和j已确定3表的name为唯一标示
						String cipinId = request
								.getParameter(i + "cipinId" + j);
						String cipinType = request.getParameter(i + "cipinType"
								+ j);
						String cipinSpecies = request.getParameter(i
								+ "cipinSpecies" + j);
						String cipinNum = request.getParameter(i + "cipinNum"
								+ j);
						Cipin cipin = new Cipin();

						Integer cipinId1 = Integer.parseInt(cipinId);
						cipin.setCipinId(cipinId1);

						// 将2表的主键值给3表外键（要不直接给3表加外键，会导致，没有外键的三表会有空的信息）
						cipin.setTrackId(workCard.getTrackId());
						cipin.setCipinType(cipinType);
						cipin.setCipinSpecies(cipinSpecies);
						cipin.setCipinNum(cipinNum);

						listCipin.add(cipin);
						trackCardSelectService.updateByCipinId(listCipin);
					}
				}
				trackCardSelectService.updateByTrackId(listWorkCard);
			}
			// 更新trackCard表的信息
			trackCardSelectService.updateByPrimaryKey(trackCard);
			return "redirect:getTrackCard";
		} else {
			// 更新trackCard表的信息
			trackCardSelectService.updateByPrimaryKey(trackCard);
			// 跟踪单的后两个表的信息android录入，所以为空时提交会报错，防止未录入时报错
			return "genzongdanlist";
		}
	}

	// 1.吴永-根据客户和产品名称、计划单号、批次号查询ProductTest表的信息
	@RequestMapping("/SelectProductTest")
	public String SelectProductTest(Integer pageNow, String client,
			String materialNo, String start_date, String end_date, Model model)
			throws Exception {
		List<ProductTest> listproductTest = new ArrayList<ProductTest>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}
		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = productTestMapper.selectProductTestByParamtotalCount(
				client, materialNo, start_date, end_date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listproductTest = productTestService.selectProductTestByParam(
				page.getStartPos(), page.getPageSize(), client, materialNo,
				start_date, end_date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listproductTest", listproductTest);
		model.addAttribute("page", page);
		model.addAttribute("client", client);
		model.addAttribute("materialNo", materialNo);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);
		// 返回到指定的jsp页面
		return "productTestList";

	}

	// 8.吴永-根据车间和操作工、图号、物料号进行生产记录查询，分页
	// 存到生产记录查询的新建pojo类ProductionRecordInquiry。
	@RequestMapping("/SelectProductionRecordInquiry")
	public String SelectProductionRecordInquiry(Integer pageNow,
			String shop_name, String operator, String client_material_no,
			String material_no, Model model) throws Exception {
		List<ProductionRecordInquiry> listproductionRecordInquiry = new ArrayList<ProductionRecordInquiry>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = trackCardMapper
				.selectProductionRecordInquiryParamtotalCount(shop_name,
						operator, client_material_no, material_no);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listproductionRecordInquiry = productTestService
				.selectProductionRecordInquiryParam(page.getStartPos(),
						page.getPageSize(), shop_name, operator,
						client_material_no, material_no);

		// controller层中使model能返回一个jsp页面
		model.addAttribute("listproductionRecordInquiry",
				listproductionRecordInquiry);
		model.addAttribute("page", page);
		model.addAttribute("shop_name", shop_name);
		model.addAttribute("operator", operator);
		model.addAttribute("client_material_no", client_material_no);
		model.addAttribute("material_no", material_no);
		// 返回到指定的jsp页面
		return "productRecord";

	}

	// 查询（根据图号，物料号，模具名称）
	// 模具出入库记录分页查询
	@RequestMapping("/moldRecord")
	public String moldRecord(Integer pageNow, String materialNo,
			String batchNo,String moldNo, Model model) {
		List<MoldRecord> listmoldrecord = new ArrayList<MoldRecord>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = moldRecordMapper.selectByPrimarytotalCount(materialNo, batchNo, moldNo);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listmoldrecord = moldRecordService.selectByPrimary(page.getStartPos(),
				page.getPageSize(), materialNo,batchNo,moldNo);
		model.addAttribute("listmoldrecord", listmoldrecord);
		model.addAttribute("page", page);
		model.addAttribute("materialNo", materialNo);
		model.addAttribute("batchNo", batchNo);		
		model.addAttribute("moldNo", moldNo);
		return "moldrecord";

	}

	// 1、领辅料审批查询功能：条件（根据物料号、起始日期、截止日期查询）分页查询
	// 领辅料审批记录分页查询
	@RequestMapping("/toApproveSecMaterial")
	public String toApproveSecMaterial(HttpSession session, Integer pageNow,
			String shopName, String start_date, String end_date, Model model)
			throws Exception {
		List<GetSecMaterials> listmaterial = new ArrayList<GetSecMaterials>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）

		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 领料审批记录分页查询返回行数totalCount
		totalCount = getSecMaterialsMapper
				.selectGetSecMaterialsByParamtotalCount(shopName, start_date,
						end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listmaterial = secMaterialService.selectGetSecMaterialsByParam(
				page.getStartPos(), page.getPageSize(), shopName, start_date,
				end_date);

		// controller层中使model能返回一个jsp页面
		model.addAttribute("listmaterial", listmaterial);
		model.addAttribute("page", page);
		model.addAttribute("shopName", shopName);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		// (更新查询数量的变化)车间主任需求：添加对审批数量的查询
		Integer num2 = getSecMaterialsMapper.selectSecGetMaterialCount();
		session.setAttribute("num2", num2);

		return "approve_sec_material";
	}

	// 2查看功能

	// 2.1查看功能（两个表通过内键getMaterialId查询）
	@RequestMapping("/lookSecMaterial")
	public String lookSecMaterial(Integer getMaterialsId, Model model)
			throws Exception {

		// 第一张表GetMaterial的查询
		GetSecMaterials getSecMaterials = secMaterialService
				.selectGetSecMaterialsKey(getMaterialsId);
		model.addAttribute("getSecMaterials", getSecMaterials);

		// 第二张表GetDetail的查询
		List<GetSecDetail> listGetSecDetail = new ArrayList<GetSecDetail>();
		listGetSecDetail = secMaterialService.selectByKey(getMaterialsId);
		model.addAttribute("listGetSecDetail", listGetSecDetail);

		return "lookSecmaterial";
	}

	// 3审批or不审批功能

	// 3.1审批功能（将已审批由0到1）

	@RequestMapping("/approveSecMaterial")
	public String approveSecMaterial(GetMaterial getMaterial,
			Integer getMaterialsId, HttpSession session, Model model)
			throws Exception {

		// 得到登录的人赋于审批的人
		Person person = (Person) session.getAttribute("user");
		String approver = person.getPersonName();

		secMaterialService.updateByKey(getMaterialsId, approver);
		return "redirect:toApproveSecMaterial.action";

	}

	// 3.2审批功能（将审批不通过由0到2）

	@RequestMapping("/notapproveSecMaterial")
	public String notapproveSecMaterial(GetMaterial getMaterial,
			Integer getMaterialsId, HttpSession session, Model model)
			throws Exception {

		// 得到登录的人赋于审批的人
		Person person = (Person) session.getAttribute("user");
		String approver = person.getPersonName();

		secMaterialService.updatenotByKey(getMaterialsId, approver);
		return "redirect:toApproveSecMaterial.action";

	}

	// 1、领料审批查询功能：条件（根据物料号、起始日期、截止日期查询）分页查询
	// 领料审批记录分页查询
	@RequestMapping("/SelectMaterial")
	public String SelectMaterial(HttpSession session, Integer pageNow,
			String material_no, String start_date, String end_date, Model model)
			throws Exception {
		List<GetMaterial> listmaterial = new ArrayList<GetMaterial>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）

		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 领料审批记录分页查询返回行数totalCount
		totalCount = getMaterialMapper.selectGetMaterialByParamtotalCount(
				material_no, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listmaterial = getMaterialService.selectGetMaterialByParam(
				page.getStartPos(), page.getPageSize(), material_no,
				start_date, end_date);

		// controller层中使model能返回一个jsp页面
		model.addAttribute("listmaterial", listmaterial);
		model.addAttribute("page", page);
		model.addAttribute("material_no", material_no);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		// (更新查询数量的变化)车间主任需求：添加对审批数量的查询
		Integer num1 = getMaterialMapper.selectGetMaterialCount();
		session.setAttribute("num1", num1);

		return "selectMaterials";
	}

	// 2查看功能

	// 2.1查看功能（两个表通过内键getMaterialId查询）
	@RequestMapping("/lookMaterial")
	public String lookMaterial(Integer getMaterialId, Model model)
			throws Exception {

		// 第一张表GetMaterial的查询
		GetMaterial getMaterial = getMaterialService
				.selectByPrimaryKey(getMaterialId);
		model.addAttribute("getMaterial", getMaterial);

		// 第二张表GetDetail的查询
		List<GetDetail> listGetDetail = new ArrayList<GetDetail>();
		listGetDetail = getMaterialService.selectByKey(getMaterialId);
		model.addAttribute("listGetDetail", listGetDetail);

		return "lookMaterials";
	}

	// 3审批or不审批功能

	// 3.1审批不通过功能
	@RequestMapping("/notapproveMaterial")
	public String notapproveMaterial(GetMaterial getMaterial,
			Integer getMaterialId, HttpSession session, Model model)
			throws Exception {
		// 得到登录的人赋于审批的人
		Person person = (Person) session.getAttribute("user");
		String approver = person.getPersonName();

		getMaterialService.updatenotByKey(getMaterialId, approver);
		return "redirect:SelectMaterial.action";
	}

	// 3.2审批功能（将是否审批由0到1）

	@RequestMapping("/approveMaterial")
	public String approve(GetMaterial getMaterial, Integer getMaterialId,
			HttpSession session, Model model) throws Exception {

		// 得到登录的人赋于审批的人
		Person person = (Person) session.getAttribute("user");
		String approver = person.getPersonName();

		getMaterialService.updateByKey(getMaterialId, approver);
		return "redirect:SelectMaterial.action";

	}

	// 1.工资查询
	// 工资单记录查询分页
	@RequestMapping("/SalarySelect")
	public String SalarySelect(Integer pageNow, String operator,
			String shop_name, String date, Model model) {
		List<WorkCard> listsalary = new ArrayList<WorkCard>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = workCardMapper.SelectByPrimarytotalCount(operator,
				shop_name, date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listsalary = salaryService.SelectByPrimary(page.getStartPos(),
				page.getPageSize(), operator, shop_name, date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listsalary", listsalary);
		model.addAttribute("page", page);
		model.addAttribute("operator", operator);
		model.addAttribute("shop_name", shop_name);
		model.addAttribute("date", date);

		// 用于回显得查询条件供excel下载用
		WorkCard HX = new WorkCard();
		if (operator != null && !("".equals(operator))) {
			HX.setOperator(operator);
		}
		if (shop_name != null && !("".equals(shop_name))) {
			HX.setShopName(shop_name);
		}
		if (date != null && !("".equals(date))) {
			HX.setProduceDate(date);
		}
		model.addAttribute("HX", HX);
		return "salary";
	}

	// 2.工资详细信息查询
	// 工资单详细记录分页
	@RequestMapping("/SalaryDetailSelect")
	public String SalaryDetailSelect(Integer pageNow, String shop_name,
			String operator, String date, Model model) {

		// 2.第一、二张表的查询
		List<Salary> listdetailsalary = new ArrayList<Salary>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = trackCardMapper.SelectByPrimaryDatetotalCount(shop_name,
				operator, date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listdetailsalary = salaryService.SelectByPrimaryDate(
				page.getStartPos(), page.getPageSize(), shop_name, operator,
				date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listdetailsalary", listdetailsalary);
		model.addAttribute("page", page);
		model.addAttribute("operator", operator);
		model.addAttribute("shop_name", shop_name);
		model.addAttribute("date", date);

		// 用于回显得查询条件供excel下载用
		Salary HX = new Salary();
		if (operator != null && !("".equals(operator))) {
			HX.setOperator(operator);
		}
		if (shop_name != null && !("".equals(shop_name))) {
			HX.setShopName(shop_name);
		}
		if (date != null && !("".equals(date))) {
			HX.setProduceDate(date);
		}
		model.addAttribute("HX", HX);

		return "salary_detail";
	}

	// 3.工资单信息的excel导出
	@RequestMapping("/Salaryexport")
	public void Salaryexport(Integer pageNow, HttpServletResponse response,
			String operator, String shop_name, String date) throws Exception {

		// 分页使用
		int totalCount = 0;
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		// 文件名
		String fileName = "工资单信息.xlsx";
		if (date == null && shop_name == null && operator == null) {

			response.setHeader("Content-disposition", "attachment;filename="
					+ new String(fileName.getBytes("gb2312"), "ISO8859-1")); // 设置文件头编码格式
			response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
			response.setHeader("Cache-Control", "no-cache");// 设置头
			response.setDateHeader("Expires", 0);// 设置日期头
		} else {
			response.setHeader("Content-disposition",
					"attachment;filename="+ date+ shop_name+ operator+ new String(fileName.getBytes("gb2312"),
									"ISO8859-1")); // 设置文件头编码格式
			response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
			response.setHeader("Cache-Control", "no-cache");// 设置头
			response.setDateHeader("Expires", 0);// 设置日期头
		}

		// 文件标题栏
		String[] cellTitle = { "操作工", "工资" };
		try {

			// 声明一个工作薄
			XSSFWorkbook workBook = null;
			workBook = new XSSFWorkbook();

			// 生成一个表格
			XSSFSheet sheet = workBook.createSheet();
			workBook.setSheetName(0, "工资单");

			// 创建表格标题行 第2行(循环将标题值赋于第1行)
			XSSFRow titleRow = sheet.createRow(1);
			for (int i = 0; i < cellTitle.length; i++) {
				titleRow.createCell(i).setCellValue(cellTitle[i]);
			}

			// 创建行
			XSSFRow row = sheet.createRow((short) 0);
			// 创建单元格
			XSSFCell cell = null;

			// 第一行标题栏
			row = sheet.createRow(0);
			cell = row.createCell(0);
			cell.setCellValue(date);
			cell = row.createCell(1);
			cell.setCellValue(shop_name);
			cell = row.createCell(2);
			cell.setCellValue(operator);
			cell = row.createCell(3);
			cell.setCellValue("工资单信息");

			// 设置居中
			XSSFCellStyle cellStyle = workBook.createCellStyle();
			cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);

			// 设置列宽
			sheet.setColumnWidth(0, 25 * 256);
			sheet.setColumnWidth(1, 20 * 256);
			sheet.setColumnWidth(3, 10 * 256);

			// 4.历史数据，业务数据，不用关注
			List<WorkCard> listsalary = new ArrayList<WorkCard>();
			listsalary = salaryService.SelectByPrimary(page.getStartPos(),
					page.getPageSize(), operator, shop_name, date);

			if (listsalary != null && listsalary.size() > 0) {

				// 5.将历史数据添加到单元格中 (先列后行)
				for (int j = 0; j < listsalary.size(); j++) {
					row = sheet.createRow(j + 2);
					cell = row.createCell(0);
					cell.setCellValue(listsalary.get(j).getOperator());
					double gzd = 0.0;

					if (listsalary.get(j).getShopName().equals("仪表工段")) {

						if (listsalary.get(j).getPrice() != null
								&& !("".equals(listsalary.get(j).getPrice()))) {
							gzd = (Double.valueOf(listsalary.get(j)
									.getHegeNum()) + Double.valueOf(listsalary
									.get(j).getTotalCipinNum()))
									* Double.valueOf(listsalary.get(j)
											.getPrice());

							cell = row.createCell(1);
							cell.setCellValue(String.valueOf(gzd));
						} else {
							cell = row.createCell(1);
							cell.setCellValue("临时工价未审批!");
						}

					} else {

						if (listsalary.get(j).getPrice() != null
								&& !("".equals(listsalary.get(j).getPrice()))) {
							gzd = Double
									.valueOf(listsalary.get(j).getHegeNum())
									* Double.valueOf(listsalary.get(j)
											.getPrice());

							cell = row.createCell(1);
							cell.setCellValue(String.valueOf(gzd));
						} else {

							cell = row.createCell(1);
							cell.setCellValue("临时工价未审批!");
						}

					}

				}
			}

			// 将文件写进输出流
			workBook.write(response.getOutputStream());
			// .flush()写出缓冲区的内容
			response.getOutputStream().flush();
			// 关闭输出流
			response.getOutputStream().close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// 3.工资单详细信息的excel导出
	@RequestMapping("/Salarydetailexport")
	public void Salarydetailexport(Integer pageNow,
			HttpServletResponse response, String operator, String shop_name,
			String date) throws Exception {
		// 分页使用
		int totalCount = 0;
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		// 文件名
		String fileName = "工资单详细信息.xlsx";
		if (date == null && shop_name == null && operator == null) {
			response.setHeader("Content-disposition", "attachment;filename="
					+ new String(fileName.getBytes("gb2312"), "ISO8859-1")); // 设置文件头编码格式
			response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
			response.setHeader("Cache-Control", "no-cache");// 设置头
			response.setDateHeader("Expires", 0);// 设置日期头
		} else {
			response.setHeader("Content-disposition",
					"attachment;filename="+ date+ shop_name+ operator
							+ new String(fileName.getBytes("gb2312"),"ISO8859-1")); // 设置文件头编码格式
			response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
			response.setHeader("Cache-Control", "no-cache");// 设置头
			response.setDateHeader("Expires", 0);// 设置日期头
		}

		// 文件标题栏
		String[] cellTitle = { "生产日期", "图号", "物料号", "车间名称", "产品规格", "工序","操作工", "合格数量", "工价", "料废工资", "工资"
				,"工废种类1","工废种类2", "工废种类3", "工废种类4", "工废种类5", "工废种类6", "工废种类7", "工废种类8", "工废种类9", "工废种类10", "工废种类11", "工废种类12"};
		try {
			// 声明一个工作薄
			XSSFWorkbook workBook = null;
			workBook = new XSSFWorkbook();

			// 生成一个表格
			XSSFSheet sheet = workBook.createSheet();
			workBook.setSheetName(0, "工资单详细信息");

			// 创建表格标题行 第一行(循环将标题值赋于第一行)
			XSSFRow titleRow = sheet.createRow(1);
			for (int i = 0; i < cellTitle.length; i++) {
				titleRow.createCell(i).setCellValue(cellTitle[i]);
			}

			// 设置居中
			XSSFCellStyle cellStyle = workBook.createCellStyle();
			cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);

			// 设置列宽
			sheet.setColumnWidth(0, 20 * 256);
			sheet.setColumnWidth(1, 20 * 256);
			sheet.setColumnWidth(2, 20 * 256);
			sheet.setColumnWidth(4, 25 * 256);
			sheet.setColumnWidth(5, 20 * 256);
			sheet.setColumnWidth(6, 25 * 256);
			sheet.setColumnWidth(9, 20 * 256);
			sheet.setColumnWidth(10, 20 * 256);
			
			sheet.setColumnWidth(11, 15 * 256);
			sheet.setColumnWidth(12, 15 * 256);
			sheet.setColumnWidth(13, 15 * 256);
			sheet.setColumnWidth(14, 15 * 256);
			sheet.setColumnWidth(15, 15 * 256);
			sheet.setColumnWidth(16, 15 * 256);
			sheet.setColumnWidth(17, 15 * 256);
			sheet.setColumnWidth(18, 15 * 256);
			sheet.setColumnWidth(19, 15 * 256);
			sheet.setColumnWidth(20, 15 * 256);
			sheet.setColumnWidth(21, 15 * 256);
			sheet.setColumnWidth(22, 15 * 256);
			// 创建行
			XSSFRow row = sheet.createRow((short) 0);
			// 创建单元格
			XSSFCell cell = null;

			// 第一行标题栏
			row = sheet.createRow(0);
			cell = row.createCell(0);
			cell.setCellValue(date);
			cell = row.createCell(1);
			cell.setCellValue(shop_name);
			cell = row.createCell(2);
			cell.setCellValue(operator);
			cell = row.createCell(3);
			cell.setCellValue("工资单详细信息");

			// 4.历史数据，业务数据，不用关注
			List<Salary> listdetailsalary = new ArrayList<Salary>();
			listdetailsalary = salaryService.SelectByPrimaryDate(page.getStartPos(), page.getPageSize(), shop_name,operator, date);
			double gzd = 0.0;
			double lfgz = 0.0;

			if (listdetailsalary != null && listdetailsalary.size() > 0) {

				// 5.将历史数据添加到单元格中 (先列后行)
				for (int j = 0; j < listdetailsalary.size(); j++) {
					row = sheet.createRow(j + 2);
					cell = row.createCell(0);
					cell.setCellValue(listdetailsalary.get(j).getProduceDate());
					cell = row.createCell(1);
					cell.setCellValue(listdetailsalary.get(j)
							.getClientMaterialNo());
					cell = row.createCell(2);
					cell.setCellValue(listdetailsalary.get(j).getMaterialNo());
					cell = row.createCell(3);
					cell.setCellValue(listdetailsalary.get(j).getShopName());
					cell = row.createCell(4);
					cell.setCellValue(listdetailsalary.get(j).getProductSpec());
					cell = row.createCell(5);
					cell.setCellValue(listdetailsalary.get(j).getProcessName());
					cell = row.createCell(6);
					cell.setCellValue(listdetailsalary.get(j).getOperator());
					
					//合格品数
					cell = row.createCell(7);
					cell.setCellValue(listdetailsalary.get(j).getHegeNum());
					
					//工价显示
					if (listdetailsalary.get(j).getPrice() != null&& (!("".equals(listdetailsalary.get(j).getPrice())))){
					cell = row.createCell(8);
					cell.setCellValue(listdetailsalary.get(j).getPrice());
					}else{
					cell = row.createCell(8);
					cell.setCellValue("无");
					}
					// 料废工资
					if (listdetailsalary.get(j).getShopName().equals("仪表工段")) {
						if (listdetailsalary.get(j).getPrice() != null
								&& (!("".equals(listdetailsalary.get(j)
										.getPrice())))) {
							double lfgz1 = Double.valueOf(listdetailsalary.get(
									j).getTotalCipinNum())
									* Double.valueOf(listdetailsalary.get(j)
											.getPrice());
							lfgz = (double) Math.round((lfgz1) * 1000) / 1000;
							cell = row.createCell(9);
							cell.setCellValue(String.valueOf(lfgz) + "");

						} else {
							cell = row.createCell(9);
							cell.setCellValue("临时工价未审批！");
						}

					} else {
						lfgz = 0.0;
						cell = row.createCell(9);
						cell.setCellValue(String.valueOf(lfgz) + "");
					}
					

					// 总工资（要考虑料废工资）
					if (listdetailsalary.get(j).getShopName().equals("仪表工段")) {
						if (listdetailsalary.get(j).getPrice() != null
								&& !("".equals(listdetailsalary.get(j)
										.getPrice()))) {
							gzd = (Double.valueOf(listdetailsalary.get(j)
									.getHegeNum()) + Double
									.valueOf(listdetailsalary.get(j)
											.getTotalCipinNum()))
									* Double.valueOf(listdetailsalary.get(j)
											.getPrice());
							cell = row.createCell(10);
							cell.setCellValue(String.valueOf(gzd) + "");
						} else {
							cell = row.createCell(10);
							cell.setCellValue("临时工价未审批!");
						}

						
					} else {
						if (listdetailsalary.get(j).getPrice() != null
								&& !("".equals(listdetailsalary.get(j)
										.getPrice()))) {
							gzd = Double.valueOf(listdetailsalary.get(j)
									.getHegeNum())
									* Double.valueOf(listdetailsalary.get(j)
											.getPrice());
							cell = row.createCell(10);
							cell.setCellValue(String.valueOf(gzd) + "");
						} else {
							cell = row.createCell(10);
							cell.setCellValue("临时工价未审批!");
						}
					}

					 //工废种类和数目
					for(int i=0;i<listdetailsalary.get(j).getCipin().size();i++){
						cell = row.createCell(11+i);
						cell.setCellValue(listdetailsalary.get(j).getCipin().get(i).getCipinSpecies()+"("
						+listdetailsalary.get(j).getCipin().get(i).getCipinNum()+")");
						
					}
				}
			}
			// 将文件写进输出流
			workBook.write(response.getOutputStream());
			// .flush()写出缓冲区的内容
			response.getOutputStream().flush();
			// 关闭输出流
			response.getOutputStream().close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	// 1.领辅料单查询//领辅料记录分页
	// 领辅料记录分页
	@RequestMapping("/toInputSecMaterialsRecord")
	public String toInputSecMaterialsRecord(Integer pageNow, String shopName,
			String start_date, String end_date, Model model) {
		// 根据物料号查询是否二表的半成品批次号存在

		// 1.1第一、二张表的查询
		List<InputSec> listinputmaterial = new ArrayList<InputSec>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = getSecMaterialsMapper.selectinputSecGetMaterialtotalCount(
				shopName, start_date, end_date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listinputmaterial = secMaterialService.selectinputSecGetMaterial(
				page.getStartPos(), page.getPageSize(), shopName, start_date,
				end_date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listinputmaterial", listinputmaterial);

		model.addAttribute("page", page);
		model.addAttribute("shopName", shopName);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		return "input_secrecord";

	}

	// 1.退辅料单查询//退辅料记录分页
	// 退辅料记录分页
	@RequestMapping("/toOutputSecMaterialsRecord")
	public String toOutputSecMaterialsRecord(Integer pageNow,
			String reshopName, String start_date, String end_date, Model model) {
		// 根据物料号查询是否二表的半成品批次号存在

		// 1.1第一、二张表的查询
		List<GetSecDetail> listoutputsecmaterial = new ArrayList<GetSecDetail>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = getSecDetailMapper.selectoutputSecGetMaterialtotalCount(
				reshopName, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listoutputsecmaterial = secMaterialService.selectoutputSecGetMaterial(
				page.getStartPos(), page.getPageSize(), reshopName, start_date,
				end_date);

		// controller层中使model能返回一个jsp页面
		model.addAttribute("listoutputsecmaterial", listoutputsecmaterial);

		model.addAttribute("page", page);
		model.addAttribute("reshopName", reshopName);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		return "output_secrecord";

	}

	// 1.领料单查询//领原材料记录分页
	// 领原材料记录分页
	@RequestMapping("/InputMaterialsRecord")
	public String InputMaterialsRecord(Integer pageNow, String material_no,
			String start_date, String end_date, Model model) {
		// 根据物料号查询是否二表的半成品批次号存在

		// 1.1第一、二张表的查询
		List<Input> listinputmaterial = new ArrayList<Input>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = getMaterialMapper.selectinputGetMaterialtotalCount(
				material_no, start_date, end_date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listinputmaterial = inputOrOutputRecordService.selectinputGetMaterial(
				page.getStartPos(), page.getPageSize(), material_no,
				start_date, end_date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listinputmaterial", listinputmaterial);

		model.addAttribute("page", page);
		model.addAttribute("material_no", material_no);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		return "inputrecord";

	}

	// 退原材料记录查询分页
	// 2.退料单查询
	@RequestMapping("/OutputMaterialsRecord")
	public String OutputMaterialsRecord(Integer pageNow, String material_no,
			String start_date, String end_date, Model model) {
		// 第一、二张表的查询
		List<Output> listoutputmaterial = new ArrayList<Output>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = getMaterialMapper.selectoutputGetMaterialtotalCount(
				material_no, start_date, end_date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listoutputmaterial = inputOrOutputRecordService
				.selectoutputGetMaterial(page.getStartPos(),
						page.getPageSize(), material_no, start_date, end_date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listoutputmaterial", listoutputmaterial);
		model.addAttribute("page", page);
		model.addAttribute("material_no", material_no);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		return "outputrecord";

	}

	// 3.领退料查看详情
	// 3.1领料详情页面
	@RequestMapping("toInputRecord")
	public String toInputRecord(Integer getMaterialId, Integer detailId,
			Model model) throws Exception {
		// 第一、二张表GetMaterial的查询
		List<Input> inputrecord = new ArrayList<Input>();
		inputrecord = inputOrOutputRecordService.selectId(getMaterialId,
				detailId);
		model.addAttribute("inputrecord", inputrecord);
		return "lookinputrecord";
	}

	// 3.2退料详情页面
	@RequestMapping("toOutputRecord")
	public String toOutputRecord(Integer getMaterialId, Integer detailId,
			Model model) throws Exception {
		// 第一、二张表GetMaterial的查询
		List<Output> outputrecord = new ArrayList<Output>();
		outputrecord = inputOrOutputRecordService.selectById(getMaterialId,
				detailId);
		model.addAttribute("outputrecord", outputrecord);
		return "lookoutputrecord";
	}

	
	// 4.领原材料外协的excel导出
		@RequestMapping("/toMaterialexcel")
		public void toMaterialexcel(Integer pageNow, HttpServletResponse response,
				String material_no, String start_date, String end_date) throws Exception {

			// 分页使用
			int totalCount = 0;
			Page page = null;
			if (pageNow != null) {
				page = new Page(totalCount, pageNow);
			} else {
				page = new Page(totalCount, 1);
			}
			

		    Date day = new Date();
			SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");

			// 文件名
			String fileName = "-原材料外协单据.xlsx";
			if (material_no == null && start_date == null && end_date == null) {

				response.setHeader("Content-disposition", "attachment;filename="
						+ new String(fileName.getBytes("gb2312"), "ISO8859-1")); // 设置文件头编码格式
				response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
				response.setHeader("Cache-Control", "no-cache");// 设置头
				response.setDateHeader("Expires", 0);// 设置日期头
			} else {
				response.setHeader("Content-disposition",
						"attachment;filename="+df.format(day)+ material_no+ new String(fileName.getBytes("gb2312"),
										"ISO8859-1")); // 设置文件头编码格式
				response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
				response.setHeader("Cache-Control", "no-cache");// 设置头
				response.setDateHeader("Expires", 0);// 设置日期头
			}

			// 文件标题栏
			String[] cellTitle = { "图号", "物料号","批次号","产品名称","产品规格","原材料批次号","材料名称","材料编号","数量","领料日期"};
			try {

				// 声明一个工作薄
				XSSFWorkbook workBook = null;
				workBook = new XSSFWorkbook();

				// 生成一个表格
				XSSFSheet sheet = workBook.createSheet();
				workBook.setSheetName(0, "原材料外协单据");

				// 创建表格标题行 第2行(循环将标题值赋于第1行)
				XSSFRow titleRow = sheet.createRow(1);
				for (int i = 0; i < cellTitle.length; i++) {
					titleRow.createCell(i).setCellValue(cellTitle[i]);
				}

				// 创建行
				XSSFRow row = sheet.createRow((short) 0);
				// 创建单元格
				XSSFCell cell = null;

				// 第一行标题栏
				row = sheet.createRow(0);
				cell = row.createCell(2);
				cell.setCellValue(material_no);
				cell = row.createCell(3);
				cell.setCellValue(df.format(day));
				cell = row.createCell(4);
				cell.setCellValue("原材料外协单据");

				// 设置居中
				XSSFCellStyle cellStyle = workBook.createCellStyle();
				cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);

				// 设置列宽
				sheet.setColumnWidth(0, 20 * 256);
				sheet.setColumnWidth(1, 20 * 256);
				sheet.setColumnWidth(2, 20 * 256);
				sheet.setColumnWidth(3, 20 * 256);
				sheet.setColumnWidth(4, 20 * 256);
				sheet.setColumnWidth(5, 25 * 256);
				sheet.setColumnWidth(6, 20 * 256);
				sheet.setColumnWidth(7, 20 * 256);
				sheet.setColumnWidth(9, 20 * 256);

				// 4.历史数据，业务数据，不用关注
				List<Input> listinputmaterial = new ArrayList<Input>();
				// 用于查询条件未输时的判断查询
				// 1、起始为空赋给一个特别小的值（1000-1-1）
				if ((start_date == null || start_date == "")
						&& ((end_date != null || end_date != ""))) {
					start_date = "1000-1-1";
				}
				// 2、截止为空赋给一个特别大的值（今天）
				if ((start_date != null || start_date != "")
						&& ((end_date == null || end_date == ""))) {
					SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
					end_date = sf.format(new Date());

				}
				// 3、起始截止都为空两个都赋值
				if ((start_date == null || start_date == "")
						&& (end_date == null || end_date == "")) {
					start_date = "1000-1-1";
					SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
					end_date = sf.format(new Date());
				}
				listinputmaterial = inputOrOutputRecordService.selectinputGetMaterial(
						page.getStartPos(), page.getPageSize(), material_no,
						start_date, end_date);
				
				if (listinputmaterial != null && listinputmaterial.size() > 0) {

					// 5.将历史数据添加到单元格中 (先列后行)
					for (int j = 0; j < listinputmaterial.size(); j++) {
						row = sheet.createRow(j + 2);
						cell = row.createCell(0);
						cell.setCellValue(listinputmaterial.get(j).getClientMaterialNo());
						cell = row.createCell(1);
						cell.setCellValue(listinputmaterial.get(j).getMaterialNo());
						cell = row.createCell(2);
						cell.setCellValue(listinputmaterial.get(j).getBatchNo());
						cell = row.createCell(3);
						cell.setCellValue(listinputmaterial.get(j).getMaterialName());
						cell = row.createCell(4);
						cell.setCellValue(listinputmaterial.get(j).getProductSpec());
						cell = row.createCell(5);
						cell.setCellValue(listinputmaterial.get(j).getMaterialBatchNo());
						cell = row.createCell(6);
						cell.setCellValue(listinputmaterial.get(j).getCailiaoMc());
						cell = row.createCell(7);
						cell.setCellValue(listinputmaterial.get(j).getCailiaoBh());
						cell = row.createCell(8);
						cell.setCellValue(listinputmaterial.get(j).getMaterialNum());
						cell = row.createCell(9);
						cell.setCellValue(listinputmaterial.get(j).getGetDate());
					}
				}

				// 将文件写进输出流
				workBook.write(response.getOutputStream());
				// .flush()写出缓冲区的内容
				response.getOutputStream().flush();
				// 关闭输出流
				response.getOutputStream().close();
			} catch (Exception e) {
				e.printStackTrace();
			}

		}
			
			
	// 1.半成品出库查询
	// 半成品出库记录分页
	@RequestMapping("/InputMiddleMaterialsRecord")
	public String InputMiddleMaterialsRecord(Integer pageNow,
			String material_no, String start_date, String end_date, Model model) {

		// 1.1第一、二张表的查询
		List<ProductRecord> listinputmaterial = new ArrayList<ProductRecord>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = productRecordMapper.selectInputMiddleRecordtotalCount(
				material_no, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listinputmaterial = inputOrOutputRecordService.selectInputMiddleRecord(
				page.getStartPos(), page.getPageSize(), material_no,
				start_date, end_date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listinputmaterial", listinputmaterial);
		model.addAttribute("page", page);
		model.addAttribute("material_no", material_no);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);
		return "inputmiddle";

	}

	// 2.半成品入库查询
	// 半成品入库记录分页
	@RequestMapping("/OutputMiddleMaterialsRecord")
	public String OutputMiddleMaterialsRecord(Integer pageNow,
			String material_no, String start_date, String end_date, Model model) {
		// 第一、二张表的查询
		List<ProductRecord> listoutputmaterial = new ArrayList<ProductRecord>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = productRecordMapper.selectOutputMiddleRecordtotalCount(
				material_no, start_date, end_date);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listoutputmaterial = inputOrOutputRecordService
				.selectOutputMiddleRecord(page.getStartPos(),
						page.getPageSize(), material_no, start_date, end_date);
		// controller层中使model能返回一个jsp页面
		model.addAttribute("listoutputmaterial", listoutputmaterial);
		model.addAttribute("page", page);
		model.addAttribute("material_no", material_no);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);

		return "outputmiddle";

	}

	// 3.领半成品查看详情
	// 3.1领半成品详情页面
	@RequestMapping("toInputMiddleRecord")
	public String toInputMiddleRecord(Integer recordId, Model model)
			throws Exception {
		// 第一、二张表GetMaterial的查询
		ProductRecord inputmiddlerecord = new ProductRecord();
		inputmiddlerecord = inputOrOutputRecordService.selectMiddleId(recordId);
		model.addAttribute("inputmiddlerecord", inputmiddlerecord);
		return "lookinputmiddle";
	}

		// 3.2退半成品详情页面
		@RequestMapping("toOutputMiddleRecord")
		public String toOutputMiddleRecord(Integer recordId, Model model)
				throws Exception {
			// 第一、二张表GetMaterial的查询
			ProductRecord outputmiddlerecord = new ProductRecord();
			outputmiddlerecord = inputOrOutputRecordService
					.selectMiddleId(recordId);
			model.addAttribute("outputmiddlerecord", outputmiddlerecord);
			return "lookoutputmiddle";
		}

	
	    
			
	// 月计划-未完成记录
	@RequestMapping("/toTotalPlanList")
	public String toTotalPlanList(Integer pageNow, String client,
			String clientMaterialNo, String start_date, String end_date,
			Model model) throws Exception {
		List<ProductionPlan> listProductionPlan = new ArrayList<ProductionPlan>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}
		int totalCount = 0;
		// 返回月计划的行数totalCount
		totalCount = productionPlanMapper.selectTotalPlanCount(client,
				clientMaterialNo, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}
		listProductionPlan = productionPlanService.selectTotalPlanByParam(
				page.getStartPos(), page.getPageSize(), client,
				clientMaterialNo, start_date, end_date);
		model.addAttribute("listProductionPlan", listProductionPlan);
		model.addAttribute("page", page);
		model.addAttribute("client", client);
		model.addAttribute("clientMaterialNo", clientMaterialNo);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);
		return "totalPlanList";
	}

	// 月计划-已完成记录
	@RequestMapping("/toFinishTotalPlanList")
	public String toFinishTotalPlanList(Integer pageNow, String client,
			String clientMaterialNo, String start_date, String end_date,
			Model model) throws Exception {
		List<ProductionPlan> listProductionPlan = new ArrayList<ProductionPlan>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}
		int totalCount = 0;
		// 返回月计划的行数totalCount
		totalCount = productionPlanMapper.selectFinishTotalPlanCount(client,
				clientMaterialNo, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}
		listProductionPlan = productionPlanService
				.selectFinishTotalPlanByParam(page.getStartPos(),
						page.getPageSize(), client, clientMaterialNo,
						start_date, end_date);
		model.addAttribute("listProductionPlan", listProductionPlan);
		model.addAttribute("page", page);
		model.addAttribute("client", client);
		model.addAttribute("clientMaterialNo", clientMaterialNo);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);
		return "totalPlanList";
	}

	// 月计划审批功能（将是否审批由0到1）

	@RequestMapping("/approveTotalPlanList")
	public String approveTotalPlanList(Integer planId, Model model)
			throws Exception {

		productionPlanService.updateByKey(planId);
		return "redirect:toTotalPlanList.action";

	}

	// 周计划记录-未完成记录
	@RequestMapping("/toProductionPlanList")
	public String toProductionPlanList(Integer pageNow, String client,
			String clientMaterialNo, String start_date, String end_date,
			Model model) throws Exception {
		List<ProductionPlan> listProductionPlan = new ArrayList<ProductionPlan>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}
		int totalCount = 0;
		// 返回周计划的行数
		totalCount = productionPlanMapper.selectProductionPlanCount(client,
				clientMaterialNo, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}
		listProductionPlan = productionPlanService.selectProductionPlanByParam(
				page.getStartPos(), page.getPageSize(), client,
				clientMaterialNo, start_date, end_date);
		model.addAttribute("listProductionPlan", listProductionPlan);
		model.addAttribute("page", page);
		model.addAttribute("client", client);
		model.addAttribute("clientMaterialNo", clientMaterialNo);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);
		// 判断是否没有未完成记录
		if (listProductionPlan.isEmpty()) {
			model.addAttribute("is_product", 0);
		} else {
			model.addAttribute("is_product", listProductionPlan.get(0)
					.getIsProduct());
		}
		return "ProductionPlanList";
	}

	// 周计划记录-已经完成记录
	@RequestMapping("/toFinishProductionPlanList")
	public String toFinishProductionPlanList(Integer pageNow, String client,
			String clientMaterialNo, String start_date, String end_date,
			Model model) throws Exception {
		List<ProductionPlan> listProductionPlan = new ArrayList<ProductionPlan>();
		// 用于查询条件未输时的判断查询
		// 1、起始为空赋给一个特别小的值（1000-1-1）
		if ((start_date == null || start_date == "")
				&& ((end_date != null || end_date != ""))) {
			start_date = "1000-1-1";
		}
		// 2、截止为空赋给一个特别大的值（今天）
		if ((start_date != null || start_date != "")
				&& ((end_date == null || end_date == ""))) {
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());

		}
		// 3、起始截止都为空两个都赋值
		if ((start_date == null || start_date == "")
				&& (end_date == null || end_date == "")) {
			start_date = "1000-1-1";
			SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
			end_date = sf.format(new Date());
		}
		int totalCount = 0;
		// 返回周计划的行数
		totalCount = productionPlanMapper.selectFinishProductionPlanCount(
				client, clientMaterialNo, start_date, end_date);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}
		listProductionPlan = productionPlanService
				.selectFinishProductionPlanByParam(page.getStartPos(),
						page.getPageSize(), client, clientMaterialNo,
						start_date, end_date);
		model.addAttribute("listProductionPlan", listProductionPlan);
		model.addAttribute("page", page);
		model.addAttribute("client", client);
		model.addAttribute("clientMaterialNo", clientMaterialNo);
		model.addAttribute("start_date", start_date);
		model.addAttribute("end_date", end_date);
		// 判断是否没有已完成记录
		if (listProductionPlan.isEmpty()) {
			model.addAttribute("is_product", 1);
		} else {
			model.addAttribute("is_product", listProductionPlan.get(0)
					.getIsProduct());
		}
		return "ProductionPlanList";
	}

	// 周计划审批功能（将是否审批由0到1）

	@RequestMapping("/approveProductionPlan")
	public String approveProductionPlan(Integer planId, Model model)
			throws Exception {

		productionPlanService.updateByKey(planId);
		return "redirect:toProductionPlanList.action";

	}

	@RequestMapping("/findProductionPlan")
	public String findProductionPlan(String orderNo, Model model)
			throws Exception {
		List<ProductionPlan> listProductionPlan = new ArrayList<ProductionPlan>();
		listProductionPlan = productionPlanMapper.selectZhouPlanByKey(orderNo);
		model.addAttribute("listProductionPlan", listProductionPlan);
		return "ProductionPlanList";
	}

	// 车间排产未完成记录分页查询 ：
	@RequestMapping("/toShopPlanList")
	public String toShopPlanList(Integer pageNow, String planNo,
			String shopName, String batchNo, Model model) throws Exception {
		List<ShopPlan> listShopPlan = new ArrayList<ShopPlan>();

		int totalCount = 0;
		// 返回月计划的行数totalCount
		totalCount = shopPlanMapper.selectShopPlanByParamtotalCount(planNo,
				shopName, batchNo);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listShopPlan = shopPlanService.selectShopPlanByParam(
				page.getStartPos(), page.getPageSize(), planNo, shopName,
				batchNo);
		model.addAttribute("listShopPlan", listShopPlan);
		model.addAttribute("page", page);
		model.addAttribute("planNo", planNo);
		model.addAttribute("shopName", shopName);
		model.addAttribute("batchNo", batchNo);
		// 判断是否没有未完成车间排产记录
		if (listShopPlan.isEmpty()) {
			model.addAttribute("is_product", 0);
		} else {
			model.addAttribute("is_product", listShopPlan.get(0).getIsProduct());
		}
		return "shopplanlist";
	}

	// 车间排产已完成记录分页查询 ：
	@RequestMapping("/toFinishShopPlanList")
	public String toFinishShopPlanList(Integer pageNow, String planNo,
			String shopName, String batchNo, Model model) throws Exception {
		List<ShopPlan> listShopPlan = new ArrayList<ShopPlan>();

		int totalCount = 0;
		// 返回月计划的行数totalCount
		totalCount = shopPlanMapper.selectFinishShopPlanByParamtotalCount(
				planNo, shopName, batchNo);
		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listShopPlan = shopPlanService.selectFinishShopPlanByParam(
				page.getStartPos(), page.getPageSize(), planNo, shopName,
				batchNo);
		model.addAttribute("listShopPlan", listShopPlan);
		model.addAttribute("page", page);
		model.addAttribute("planNo", planNo);
		model.addAttribute("shopName", shopName);
		model.addAttribute("batchNo", batchNo);
		// 判断是否没有已完成车间排产记录
		if (listShopPlan.isEmpty()) {
			model.addAttribute("is_product", 1);
		} else {
			model.addAttribute("is_product", listShopPlan.get(0).getIsProduct());
		}
		return "shopplanlist";
	}

	// 车间排产计划审批功能（将是否审批由0到1）

	@RequestMapping("/approveShopPlanList")
	public String approveShopPlanList(Integer planId, Model model)
			throws Exception {

		shopPlanService.updateByKey(planId);
		return "redirect:toShopPlanList.action";

	}

	// 日检记录查询
	@RequestMapping("/DailyCheck")
	public String DailyCheck(Integer pageNow, String batchNo,
			String processName, String assetNo, Model model) throws Exception {
		List<DailyCheck> listDailyCheck = new ArrayList<DailyCheck>();

		int totalCount = 0;
		// 返回查询的行数totalCount

		totalCount = dailycheckMapper.selectByPrimarytotalCount(batchNo,
				processName, assetNo);

		Page page = null;
		if (pageNow != null) {
			page = new Page(totalCount, pageNow);
		} else {
			page = new Page(totalCount, 1);
		}

		listDailyCheck = dailyCheckService.selectByPrimary(page.getStartPos(),
				page.getPageSize(), batchNo, processName, assetNo);
		model.addAttribute("listDailyCheck", listDailyCheck);
		model.addAttribute("page", page);
		model.addAttribute("batchNo", batchNo);
		model.addAttribute("processName", processName);
		model.addAttribute("assetNo", assetNo);
		return "dailycheck";
	}

	// 日检记录查询
	@RequestMapping("/lookCheck")
	public String lookCheck(Integer checkId, Model model) throws Exception {

		// 第一张表DailyCheck的查询
		DailyCheck dailyCheck = dailyCheckService.selectByPrimaryKey(checkId);
		model.addAttribute("dailyCheck", dailyCheck);

		// 第二张表CheckRecord的查询
		List<CheckRecord> listCheckRecord = new ArrayList<CheckRecord>();
		listCheckRecord = dailyCheckService.selectCheckRecordByKey(checkId);
		model.addAttribute("listCheckRecord", listCheckRecord);

		return "lookCheck";
	}

	
		// 1.部分外协记录
		@RequestMapping("/toExterAssociation")
		public String toExterAssociation(Integer pageNow, HttpServletResponse response,
				String material_no, String start_date, String end_date,Model model) throws Exception {
			List<ExterAssociation> listExterAssociation=new ArrayList<ExterAssociation>();
			
			int totalCount = 0;
			// 返回月计划的行数totalCount
			totalCount = shopTransitionMapper.selectExterAssociationtotalCount(material_no, start_date, end_date);
			Page page = null;
			if (pageNow != null) {
				page = new Page(totalCount, pageNow);
			} else {
				page = new Page(totalCount, 1);
			}

			listExterAssociation=shopPlanService.selectExterAssociation(page.getStartPos(), page.getPageSize(), material_no, start_date, end_date);
			
			model.addAttribute("listExterAssociation", listExterAssociation);
			model.addAttribute("page", page);
			model.addAttribute("material_no", material_no);
			model.addAttribute("start_date", start_date);
			model.addAttribute("end_date", end_date);
			
			return "exter_association";
		}

		
	
		// 2.部分外协的excel导出
		@RequestMapping("/toExterAssociationexcel")
		public void toExterAssociationexcel(Integer pageNow, HttpServletResponse response,
				String material_no, String start_date, String end_date) throws Exception {

			// 分页使用
			int totalCount = 0;
			Page page = null;
			if (pageNow != null) {
				page = new Page(totalCount, pageNow);
			} else {
				page = new Page(totalCount, 1);
			}
			

		    Date day = new Date();
			SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");

			// 文件名
			String fileName = "-部分外协单据.xlsx";
			if (material_no == null && start_date == null && end_date == null) {

				response.setHeader("Content-disposition", "attachment;filename="
						+ new String(fileName.getBytes("gb2312"), "ISO8859-1")); // 设置文件头编码格式
				response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
				response.setHeader("Cache-Control", "no-cache");// 设置头
				response.setDateHeader("Expires", 0);// 设置日期头
			} else {
				response.setHeader("Content-disposition",
						"attachment;filename="+df.format(day)+ "-"+material_no+ new String(fileName.getBytes("gb2312"),
										"ISO8859-1")); // 设置文件头编码格式
				response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
				response.setHeader("Cache-Control", "no-cache");// 设置头
				response.setDateHeader("Expires", 0);// 设置日期头
			}

			// 文件标题栏
			String[] cellTitle = {"计划单号","图号","物料号","批次号","产品名称","提供工段","接收工段","外协日期","外协数量"};
			try {

				// 声明一个工作薄
				XSSFWorkbook workBook = null;
				workBook = new XSSFWorkbook();

				// 生成一个表格
				XSSFSheet sheet = workBook.createSheet();
				workBook.setSheetName(0, "部分外协清单");

				// 创建表格标题行 第2行(循环将标题值赋于第1行)
				XSSFRow titleRow = sheet.createRow(1);
				for (int i = 0; i < cellTitle.length; i++) {
					titleRow.createCell(i).setCellValue(cellTitle[i]);
				}

				// 创建行
				XSSFRow row = sheet.createRow((short) 0);
				// 创建单元格
				XSSFCell cell = null;

				// 第一行标题栏
				row = sheet.createRow(0);
				cell = row.createCell(2);
				cell.setCellValue(material_no);
				cell = row.createCell(3);
				cell.setCellValue(df.format(day));
				cell = row.createCell(4);
				cell.setCellValue("部分外协单据");

				// 设置居中
				XSSFCellStyle cellStyle = workBook.createCellStyle();
				cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);

				// 设置列宽
				sheet.setColumnWidth(0, 20 * 256);
				sheet.setColumnWidth(1, 20 * 256);
				sheet.setColumnWidth(2, 20 * 256);
				sheet.setColumnWidth(3, 20 * 256);
				sheet.setColumnWidth(4, 20 * 256);
				sheet.setColumnWidth(7, 20 * 256);
				
				// 4.历史数据，业务数据，不用关注
			List<ExterAssociation> listExterAssociation=new ArrayList<ExterAssociation>();
			listExterAssociation=shopPlanService.selectExterAssociation(page.getStartPos(), page.getPageSize(), material_no, start_date, end_date);
				
			if (listExterAssociation != null && listExterAssociation.size() > 0) {

					// 5.将历史数据添加到单元格中 (先列后行)
					for (int j = 0; j < listExterAssociation.size(); j++) {
						row = sheet.createRow(j + 2);
						cell = row.createCell(0);
						cell.setCellValue(listExterAssociation.get(j).getPlanNo());
						cell = row.createCell(1);
						cell.setCellValue(listExterAssociation.get(j).getClientMaterialNo());
						cell = row.createCell(2);
						cell.setCellValue(listExterAssociation.get(j).getMaterialNo());
						cell = row.createCell(3);
						cell.setCellValue(listExterAssociation.get(j).getBatchNo());
						cell = row.createCell(4);
						cell.setCellValue(listExterAssociation.get(j).getProductName());
						cell = row.createCell(5);
						cell.setCellValue(listExterAssociation.get(j).getShop1());
						cell = row.createCell(6);
						cell.setCellValue(listExterAssociation.get(j).getShop2());
						cell = row.createCell(7);
						cell.setCellValue(listExterAssociation.get(j).getTranDate());
						cell = row.createCell(8);
						cell.setCellValue(listExterAssociation.get(j).getQualifiedNum());
						
					}
				}

				// 将文件写进输出流
				workBook.write(response.getOutputStream());
				// .flush()写出缓冲区的内容
				response.getOutputStream().flush();
				// 关闭输出流
				response.getOutputStream().close();
			} catch (Exception e) {
				e.printStackTrace();
			}

		}
		
		
	
		// 原材料外协出库记录分页
		@RequestMapping("/toMaterialsAssociation")
		public String toMaterialsAssociation(Integer pageNow, String material_no,
				String start_date, String end_date, Model model) {
			// 根据物料号查询是否二表的半成品批次号存在

			// 1.1第一、二张表的查询
			List<Input> listmaterialassociation = new ArrayList<Input>();
			// 用于查询条件未输时的判断查询
			// 1、起始为空赋给一个特别小的值（1000-1-1）
			if ((start_date == null || start_date == "")
					&& ((end_date != null || end_date != ""))) {
				start_date = "1000-1-1";
			}
			// 2、截止为空赋给一个特别大的值（今天）
			if ((start_date != null || start_date != "")
					&& ((end_date == null || end_date == ""))) {
				SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
				end_date = sf.format(new Date());

			}
			// 3、起始截止都为空两个都赋值
			if ((start_date == null || start_date == "")
					&& (end_date == null || end_date == "")) {
				start_date = "1000-1-1";
				SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
				end_date = sf.format(new Date());
			}

			int totalCount = 0;
			// 返回查询的行数totalCount

			totalCount = getMaterialMapper.selectMaterialAssociationtotalCount(material_no, start_date, end_date);

			Page page = null;
			if (pageNow != null) {
				page = new Page(totalCount, pageNow);
			} else {
				page = new Page(totalCount, 1);
			}

			listmaterialassociation = inputOrOutputRecordService.selectMaterialAssociation(page.getStartPos(), page.getPageSize(), material_no, start_date, end_date);
			// controller层中使model能返回一个jsp页面
			model.addAttribute("listmaterialassociation", listmaterialassociation);

			model.addAttribute("page", page);
			model.addAttribute("material_no", material_no);
			model.addAttribute("start_date", start_date);
			model.addAttribute("end_date", end_date);

			return "output_material_association";

		}
				
		 // 2.原材料外协出库excel导出
			@RequestMapping("/toMaterialsAssociationexcel")
			public void toMaterialsAssociationexcel(Integer pageNow, HttpServletResponse response,
					String material_no, String start_date, String end_date) throws Exception {

				// 分页使用
				int totalCount = 0;
				Page page = null;
				if (pageNow != null) {
					page = new Page(totalCount, pageNow);
				} else {
					page = new Page(totalCount, 1);
				}
				

			    Date day = new Date();
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");

				// 文件名
				String fileName = "-原材料外协出库单据.xlsx";
				if (material_no == null && start_date == null && end_date == null) {

					response.setHeader("Content-disposition", "attachment;filename="
							+ new String(fileName.getBytes("gb2312"), "ISO8859-1")); // 设置文件头编码格式
					response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
					response.setHeader("Cache-Control", "no-cache");// 设置头
					response.setDateHeader("Expires", 0);// 设置日期头
				} else {
					response.setHeader("Content-disposition",
							"attachment;filename="+df.format(day)+ "-"+material_no+ new String(fileName.getBytes("gb2312"),
											"ISO8859-1")); // 设置文件头编码格式
					response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
					response.setHeader("Cache-Control", "no-cache");// 设置头
					response.setDateHeader("Expires", 0);// 设置日期头
				}

				// 文件标题栏
				String[] cellTitle = {"图号","物料号","计划单号","批次号","产品名称","产品规格","原材料批次号","材料名称","材料编号","外协出库数量","单位","外协出库日期"};
				try {

					// 声明一个工作薄
					XSSFWorkbook workBook = null;
					workBook = new XSSFWorkbook();

					// 生成一个表格
					XSSFSheet sheet = workBook.createSheet();
					workBook.setSheetName(0, "原材料外协出库清单");

					// 创建表格标题行 第2行(循环将标题值赋于第1行)
					XSSFRow titleRow = sheet.createRow(1);
					for (int i = 0; i < cellTitle.length; i++) {
						titleRow.createCell(i).setCellValue(cellTitle[i]);
					}

					// 创建行
					XSSFRow row = sheet.createRow((short) 0);
					// 创建单元格
					XSSFCell cell = null;

					// 第一行标题栏
					row = sheet.createRow(0);
					cell = row.createCell(2);
					cell.setCellValue(material_no);
					cell = row.createCell(3);
					cell.setCellValue(df.format(day));
					cell = row.createCell(4);
					cell.setCellValue("原材料外协出库单据");

					// 设置居中
					XSSFCellStyle cellStyle = workBook.createCellStyle();
					cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);

					// 设置列宽
					sheet.setColumnWidth(0, 20 * 256);
					sheet.setColumnWidth(1, 20 * 256);
					sheet.setColumnWidth(2, 20 * 256);
					sheet.setColumnWidth(3, 20 * 256);
					sheet.setColumnWidth(4, 20 * 256);
					sheet.setColumnWidth(5, 20 * 256);
					sheet.setColumnWidth(6, 20 * 256);
					sheet.setColumnWidth(7, 20 * 256);
					sheet.setColumnWidth(8, 20 * 256);
					sheet.setColumnWidth(9, 15 * 256);
					sheet.setColumnWidth(11, 15 * 256);
					
					// 4.历史数据，业务数据，不用关注
				List<Input> listmaterialassociation=new ArrayList<Input>();
				listmaterialassociation = inputOrOutputRecordService.selectMaterialAssociation(page.getStartPos(), page.getPageSize(), material_no, start_date, end_date);
					
				if (listmaterialassociation != null && listmaterialassociation.size() > 0) {

						// 5.将历史数据添加到单元格中 (先列后行)
						for (int j = 0; j < listmaterialassociation.size(); j++) {
							row = sheet.createRow(j + 2);
							cell = row.createCell(0);
							cell.setCellValue(listmaterialassociation.get(j).getClientMaterialNo());
							cell = row.createCell(1);
							cell.setCellValue(listmaterialassociation.get(j).getMaterialNo());
							cell = row.createCell(2);
							cell.setCellValue(listmaterialassociation.get(j).getPlanNo());
							cell = row.createCell(3);
							cell.setCellValue(listmaterialassociation.get(j).getBatchNo());
							//产品名称用unqualified不合格数前端充当
							cell = row.createCell(4);
							cell.setCellValue(listmaterialassociation.get(j).getMaterialName());
							cell = row.createCell(5);
							cell.setCellValue(listmaterialassociation.get(j).getProductSpec());
							cell = row.createCell(6);
							cell.setCellValue(listmaterialassociation.get(j).getMaterialBatchNo());
							cell = row.createCell(7);
							cell.setCellValue(listmaterialassociation.get(j).getCailiaoMc());
							cell = row.createCell(8);
							cell.setCellValue(listmaterialassociation.get(j).getCailiaoBh());
							cell = row.createCell(9);
							cell.setCellValue(listmaterialassociation.get(j).getMaterialNum());
							cell = row.createCell(10);
							cell.setCellValue(listmaterialassociation.get(j).getUnit());
							cell = row.createCell(11);
							cell.setCellValue(listmaterialassociation.get(j).getGetDate());
							
						}
					}

					// 将文件写进输出流
					workBook.write(response.getOutputStream());
					// .flush()写出缓冲区的内容
					response.getOutputStream().flush();
					// 关闭输出流
					response.getOutputStream().close();
				} catch (Exception e) {
					e.printStackTrace();
				}

			}
			
			
			// 原材料外协入库记录分页
			@RequestMapping("/toIntputMaterialsAssociation")
			public String toIntputMaterialsAssociation(Integer pageNow, String material_no,
					String start_date, String end_date, Model model) {
				// 第一、二张表的查询
				List<InputMaterialAssociation> listintputmaterialassociation = new ArrayList<InputMaterialAssociation>();
				// 用于查询条件未输时的判断查询
				// 1、起始为空赋给一个特别小的值（1000-1-1）
				if ((start_date == null || start_date == "")
						&& ((end_date != null || end_date != ""))) {
					start_date = "1000-1-1";
				}
				// 2、截止为空赋给一个特别大的值（今天）
				if ((start_date != null || start_date != "")
						&& ((end_date == null || end_date == ""))) {
					SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
					end_date = sf.format(new Date());

				}
				// 3、起始截止都为空两个都赋值
				if ((start_date == null || start_date == "")
						&& (end_date == null || end_date == "")) {
					start_date = "1000-1-1";
					SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd");
					end_date = sf.format(new Date());
				}

				int totalCount = 0;
				// 返回查询的行数totalCount

				totalCount = productRecordMapper.selectIntputMaterialAssociationtotalCount(material_no, start_date, end_date);

				Page page = null;
				if (pageNow != null) {
					page = new Page(totalCount, pageNow);
				} else {
					page = new Page(totalCount, 1);
				}

				listintputmaterialassociation = inputOrOutputRecordService.selectIntputMaterialAssociation(page.getStartPos(),page.getPageSize(), material_no, start_date, end_date);
						
				// controller层中使model能返回一个jsp页面
				model.addAttribute("listintputmaterialassociation", listintputmaterialassociation);
				model.addAttribute("page", page);
				model.addAttribute("material_no", material_no);
				model.addAttribute("start_date", start_date);
				model.addAttribute("end_date", end_date);

				return "input_material_association";

			}
			
			
			 // 2.原材料外协入库excel导出
			@RequestMapping("/toInputMaterialsAssociationexcel")
			public void toInputMaterialsAssociationexcel(Integer pageNow, HttpServletResponse response,
					String material_no, String start_date, String end_date) throws Exception {

				// 分页使用
				int totalCount = 0;
				Page page = null;
				if (pageNow != null) {
					page = new Page(totalCount, pageNow);
				} else {
					page = new Page(totalCount, 1);
				}
				

			    Date day = new Date();
				SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");

				// 文件名
				String fileName = "-原材料外协入库单据.xlsx";
				if (material_no == null && start_date == null && end_date == null) {

					response.setHeader("Content-disposition", "attachment;filename="
							+ new String(fileName.getBytes("gb2312"), "ISO8859-1")); // 设置文件头编码格式
					response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
					response.setHeader("Cache-Control", "no-cache");// 设置头
					response.setDateHeader("Expires", 0);// 设置日期头
				} else {
					response.setHeader("Content-disposition",
							"attachment;filename="+df.format(day)+ "-"+material_no+ new String(fileName.getBytes("gb2312"),
											"ISO8859-1")); // 设置文件头编码格式
					response.setContentType("APPLICATION/OCTET-STREAM;charset=UTF-8");// 设置类型
					response.setHeader("Cache-Control", "no-cache");// 设置头
					response.setDateHeader("Expires", 0);// 设置日期头
				}

				// 文件标题栏
				String[] cellTitle = {"图号","物料号","计划单号","批次号","产品名称","产品规格","外协入库数量","单位","外协入库日期"};
				try {

					// 声明一个工作薄
					XSSFWorkbook workBook = null;
					workBook = new XSSFWorkbook();

					// 生成一个表格
					XSSFSheet sheet = workBook.createSheet();
					workBook.setSheetName(0, "原材料外协入库清单");

					// 创建表格标题行 第2行(循环将标题值赋于第1行)
					XSSFRow titleRow = sheet.createRow(1);
					for (int i = 0; i < cellTitle.length; i++) {
						titleRow.createCell(i).setCellValue(cellTitle[i]);
					}

					// 创建行
					XSSFRow row = sheet.createRow((short) 0);
					// 创建单元格
					XSSFCell cell = null;

					// 第一行标题栏
					row = sheet.createRow(0);
					cell = row.createCell(2);
					cell.setCellValue(material_no);
					cell = row.createCell(3);
					cell.setCellValue(df.format(day));
					cell = row.createCell(4);
					cell.setCellValue("原材料外协入库单据");

					// 设置居中
					XSSFCellStyle cellStyle = workBook.createCellStyle();
					cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);

					// 设置列宽
					sheet.setColumnWidth(0, 20 * 256);
					sheet.setColumnWidth(1, 20 * 256);
					sheet.setColumnWidth(2, 20 * 256);
					sheet.setColumnWidth(3, 20 * 256);
					sheet.setColumnWidth(4, 20 * 256);
					sheet.setColumnWidth(5, 20 * 256);
					sheet.setColumnWidth(6, 15 * 256);
					sheet.setColumnWidth(8, 15 * 256);
					
					
					// 4.历史数据，业务数据，不用关注
				List<InputMaterialAssociation> listintputmaterialassociation=new ArrayList<InputMaterialAssociation>();
				listintputmaterialassociation = inputOrOutputRecordService.selectIntputMaterialAssociation(page.getStartPos(),page.getPageSize(), material_no, start_date, end_date);
				if (listintputmaterialassociation != null && listintputmaterialassociation.size() > 0) {

						// 5.将历史数据添加到单元格中 (先列后行)
						for (int j = 0; j < listintputmaterialassociation.size(); j++) {
							row = sheet.createRow(j + 2);
							cell = row.createCell(0);
							cell.setCellValue(listintputmaterialassociation.get(j).getClientMaterialNo());
							cell = row.createCell(1);
							cell.setCellValue(listintputmaterialassociation.get(j).getMaterialNo());
							cell = row.createCell(2);
							cell.setCellValue(listintputmaterialassociation.get(j).getPlanNo());
							cell = row.createCell(3);
							cell.setCellValue(listintputmaterialassociation.get(j).getBatchNo());
							cell = row.createCell(4);
							cell.setCellValue(listintputmaterialassociation.get(j).getProductMc());
							cell = row.createCell(5);
							cell.setCellValue(listintputmaterialassociation.get(j).getProductSpec());
							cell = row.createCell(6);
							cell.setCellValue(listintputmaterialassociation.get(j).getProductNum());
							cell = row.createCell(7);
							cell.setCellValue(listintputmaterialassociation.get(j).getUnit());
							cell = row.createCell(8);
							cell.setCellValue(listintputmaterialassociation.get(j).getTransDate());
							
							
						}
					}

					// 将文件写进输出流
					workBook.write(response.getOutputStream());
					// .flush()写出缓冲区的内容
					response.getOutputStream().flush();
					// 关闭输出流
					response.getOutputStream().close();
				} catch (Exception e) {
					e.printStackTrace();
				}

			}
			
		

			 // 领原材料T+系统的选择模板并选择模板excel导出
			@RequestMapping("/importRawMaterialExcel")
			@ResponseBody
			public String importRawMaterialExcel(
					HttpServletRequest request,HttpServletResponse response,
					@RequestParam(value = "excelPath", required = false) MultipartFile excelPath)
					throws Exception {

				// 获取当天日期
				Date day = new Date();
				SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
				SimpleDateFormat df2 = new SimpleDateFormat("yyyy-MM");
				
				try {
					Workbook book1 = Workbook.getWorkbook(excelPath.getInputStream());// 获取流(会自动识别流中文件的实际地址，取得excel文件中的内容)
					String path = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
					String url=path+df1.format(day)+"原材料出库单.xls";
					//String url="C:/Users/Administrator/Desktop/"+df1.format(day)+"原材料出库单.xls";
					WritableWorkbook book= Workbook.createWorkbook(new File(url),book1);// 在原模板数据源的excel上追加数据
					WritableSheet sheet=book.getSheet(0);
					
				
					// 居中
					WritableCellFormat wf = new WritableCellFormat();
					wf.setAlignment(Alignment.CENTRE);
					
						
					// 判断原模板sheet1的名称是不是设备表，以确定导入的是不是材料出库单
					if (sheet.getName().equals("材料出库单") == false) {
						return "2";
					}
					
					for (int i = 1; i < sheet.getRows(); i++) {
						// 有效数据的行数
						Cell[] cell = sheet.getRow(i);
						// 判断有效数据的行中是否存在空
						if (cell.length == 0)
							continue;

						List<Input> listrowinMaterial=new ArrayList<Input>();
						
						
						//原材料的记录选择（除了外协的原材料出库）
						listrowinMaterial = inputOrOutputRecordService.selectRowinMaterial();
						

						
						if(listrowinMaterial != null && listrowinMaterial.size() > 0){
									// 5.将历史数据添加到单元格中 (先列后行)
									for (int j = 0; j < listrowinMaterial.size(); j++) {
										//1.单据日期
										sheet.addCell((WritableCell) new Label(0, j + 2,df1.format(day), wf));
										
										//2.单据编号
										//（需要先查询数据库中该月中原材料的单据日期数量）
										Integer num=0;
										String num1=tJdbcService.selectRowMaterialNum();
										if(num1!=null&&"".equals(num1)){
										    num=Integer.valueOf(num1);
										}else{
											num=0;
										}
										
										//判断如果是1位数拼接为：当前日期到月份"yyyy-MM-+000+num"
										if(num>=0&&num<10){
											num=num+1+j;
											sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-000"+num, wf));
										}
										//判断如果2位数拼接为：当前日期到月份"yyyy-MM-+00+num"
										if(num>=10&&num<100){
											sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-00"+num, wf));
										}
										//判断如果3位数拼接为：当前日期到月份"yyyy-MM-+0+num"
										if(num>=100){
											sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-0"+num+j, wf));
										}
							
										//3.业务类型编码（领料默认为41，退料默认为42）
										sheet.addCell((WritableCell) new Label(2, j + 2,41+"", wf));
										//4.生产车间
										sheet.addCell((WritableCell) new Label(7, j + 2,listrowinMaterial.get(j).getShopName(), wf));
										//5.仓库编码
										sheet.addCell((WritableCell) new Label(11, j + 2,"01", wf));
										//6.材料编码
										sheet.addCell((WritableCell) new Label(15, j + 2,listrowinMaterial.get(j).getCailiaoBh(), wf));
										//7.计量单位
										sheet.addCell((WritableCell) new Label(18, j + 2,listrowinMaterial.get(j).getUnit(), wf));
										//8.数量
										sheet.addCell((WritableCell) new Label(19, j + 2,listrowinMaterial.get(j).getMaterialNum(), wf));
										//9.批号(对应于我们的原材料批次号)
										sheet.addCell((WritableCell) new Label(20, j + 2,listrowinMaterial.get(j).getMaterialBatchNo(), wf));
										//10.领用人
										sheet.addCell((WritableCell) new Label(23, j + 2,listrowinMaterial.get(j).getAcceptor(), wf));
										
									}
						}
							
					}
					// 6.写入excel并关闭
					book.write();
					book.close();
					return "1";
				} catch (Exception e) {
					return "0";

				}

			}
			
		
			 // 退原材料T+系统的选择模板并选择模板excel导出
			@RequestMapping("/outportRawMaterialExcel")
			@ResponseBody
			public String outportRawMaterialExcel(
					HttpServletRequest request,HttpServletResponse response,
					@RequestParam(value = "excelPath", required = false) MultipartFile excelPath)
					throws Exception {

				// 获取当天日期
				Date day = new Date();
				SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
				SimpleDateFormat df2 = new SimpleDateFormat("yyyy-MM");
				
				try {
					Workbook book1 = Workbook.getWorkbook(excelPath.getInputStream());// 获取流(会自动识别流中文件的实际地址，取得excel文件中的内容)
					String path = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
					//String url="C:/Users/Administrator/Desktop/"+df1.format(day)+"退原材料单.xls";
					String url=path+df1.format(day)+"退原材料单.xls";
					WritableWorkbook book= Workbook.createWorkbook(new File(url),book1);// 在原模板数据源的excel上追加数据
					WritableSheet sheet=book.getSheet(0);
				
					// 居中
					WritableCellFormat wf = new WritableCellFormat();
					wf.setAlignment(Alignment.CENTRE);
					
						
					// 判断原模板sheet1的名称是不是设备表，以确定导入的是不是材料出库单
					if (sheet.getName().equals("材料出库单") == false) {
						return "2";
					}
					
					for (int i = 1; i < sheet.getRows(); i++) {
						// 有效数据的行数
						Cell[] cell = sheet.getRow(i);
						// 判断有效数据的行中是否存在空
						if (cell.length == 0)
							continue;

						List<Output> listrowoutMaterial=new ArrayList<Output>();
					
						//原材料的记录选择（除了外协的原材料出库）
						listrowoutMaterial = inputOrOutputRecordService.selectRowoutMaterial();
						
						
						if(listrowoutMaterial != null && listrowoutMaterial.size() > 0){
									// 5.将历史数据添加到单元格中 (先列后行)
									for (int j = 0; j < listrowoutMaterial.size(); j++) {
										//1.单据日期
										sheet.addCell((WritableCell) new Label(0, j + 2,df1.format(day), wf));
										
										//2.单据编号
										//（需要先查询数据库中该月中原材料的单据日期数量）
										Integer num=0;
										String num1=tJdbcService.selectRowMaterialNum();
										if(num1!=null&&"".equals(num1)){
										    num=Integer.valueOf(num1);
										}else{
											num=0;
										}
										
										//判断如果是1位数拼接为：当前日期到月份"yyyy-MM-+000+num"
										if(num>=0&&num<10){
											num=num+1+j;
											sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-000"+num, wf));
										}
										//判断如果2位数拼接为：当前日期到月份"yyyy-MM-+00+num"
										if(num>=10&&num<100){
											sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-00"+num, wf));
										}
										//判断如果3位数拼接为：当前日期到月份"yyyy-MM-+0+num"
										if(num>=100){
											sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-0"+num+j, wf));
										}
							
										//3.业务类型编码（领料默认为41，退料默认为42）
										sheet.addCell((WritableCell) new Label(2, j + 2,"42", wf));
										//4.生产车间
										sheet.addCell((WritableCell) new Label(7, j + 2,listrowoutMaterial.get(j).getShopName(), wf));
										//5.仓库编码
										sheet.addCell((WritableCell) new Label(11, j + 2,"01", wf));
										//6.材料编码
										sheet.addCell((WritableCell) new Label(15, j + 2,listrowoutMaterial.get(j).getCailiaoBh(), wf));
										//7.计量单位
										sheet.addCell((WritableCell) new Label(18, j + 2,listrowoutMaterial.get(j).getUnit(), wf));
										//8.数量
										sheet.addCell((WritableCell) new Label(19, j + 2,listrowoutMaterial.get(j).getMaterialNum(), wf));
										//9.批号(对应于我们的原材料批次号)
										sheet.addCell((WritableCell) new Label(20, j + 2,listrowoutMaterial.get(j).getMaterialBatchNo(), wf));
										//10.领用人
										sheet.addCell((WritableCell) new Label(23, j + 2,listrowoutMaterial.get(j).getAcceptor(), wf));
										
									}
						}
							
					}
					// 6.写入excel并关闭
					book.write();
					book.close();
					return "1";
				} catch (Exception e) {
					return "0";

				}

			}
			
			
			
		// 领辅料T+系统的选择模板并选择模板excel导出
		@RequestMapping("/importSecMaterialExcel")
		@ResponseBody
		public String importSecMaterialExcel(
				HttpServletRequest request,HttpServletResponse response,
				@RequestParam(value = "excelPath", required = false) MultipartFile excelPath)
				throws Exception {
			
			
			// 获取当天日期
			Date day = new Date();
			SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
			SimpleDateFormat df2 = new SimpleDateFormat("yyyy-MM");
			
			try {
				Workbook book1 = Workbook.getWorkbook(excelPath.getInputStream());// 获取流(会自动识别流中文件的实际地址，取得excel文件中的内容)
				String path = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
				String url=path+df1.format(day)+"辅料出库单.xls";
				//String url="C:/Users/Administrator/Desktop/"+df1.format(day)+"辅料出库单.xls";
				WritableWorkbook book= Workbook.createWorkbook(new File(url),book1);// 在原模板数据源的excel上追加数据
				WritableSheet sheet=book.getSheet(0);
				
				// 居中
				WritableCellFormat wf = new WritableCellFormat();
				wf.setAlignment(Alignment.CENTRE);
				
					
				// 判断原模板sheet1的名称是不是设备表，以确定导入的是不是材料出库单
				if (sheet.getName().equals("材料出库单") == false) {
					return "2";
				}
				
				for (int i = 1; i < sheet.getRows(); i++) {
					// 有效数据的行数
					Cell[] cell = sheet.getRow(i);
					// 判断有效数据的行中是否存在空
					if (cell.length == 0)
						continue;

					List<InputSec> listsecinMaterial=new ArrayList<InputSec>();
				
					//原材料的记录选择（除了外协的原材料出库）
					listsecinMaterial = secMaterialService.selectexcelinputSecGetMaterial();
					
					if(listsecinMaterial != null && listsecinMaterial.size() > 0){
								// 5.将历史数据添加到单元格中 (先列后行)
								for (int j = 0; j < listsecinMaterial.size(); j++) {
									//1.单据日期
									sheet.addCell((WritableCell) new Label(0, j + 2,df1.format(day), wf));
									
									//2.单据编号
									//（需要先查询数据库中该月中原材料的单据日期数量）
									Integer num=0;
									String num1=tJdbcService.selectRowMaterialNum();
									if(num1!=null&&"".equals(num1)){
									    num=Integer.valueOf(num1);
									}else{
										num=0;
									}
									
									//判断如果是1位数拼接为：当前日期到月份"yyyy-MM-+000+num"
									if(num>=0&&num<10){
										num=num+1+j;
										sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-000"+num, wf));
									}
									//判断如果2位数拼接为：当前日期到月份"yyyy-MM-+00+num"
									if(num>=10&&num<100){
										sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-00"+num, wf));
									}
									//判断如果3位数拼接为：当前日期到月份"yyyy-MM-+0+num"
									if(num>=100){
										sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-0"+num+j, wf));
									}
						
									//3.业务类型编码（领料默认为41，退料默认为42）
									sheet.addCell((WritableCell) new Label(2, j + 2,41+"", wf));
									//4.生产车间
									sheet.addCell((WritableCell) new Label(7, j + 2,listsecinMaterial.get(j).getShopName(), wf));
									//5.仓库编码
									sheet.addCell((WritableCell) new Label(11, j + 2,"03", wf));
									//6.材料编码
									sheet.addCell((WritableCell) new Label(15, j + 2,listsecinMaterial.get(j).getSecMaterialNo(), wf));
									//7.计量单位
									sheet.addCell((WritableCell) new Label(18, j + 2,listsecinMaterial.get(j).getUnit(), wf));
									//8.数量
									sheet.addCell((WritableCell) new Label(19, j + 2,listsecinMaterial.get(j).getNum(), wf));
									//10.领用人
									sheet.addCell((WritableCell) new Label(23, j + 2,listsecinMaterial.get(j).getAcceptor(), wf));
									
								}
					}
						
				}
				// 6.写入excel并关闭
				book.write();
				book.close();
				return "1";
			} catch (Exception e) {
				return "0";

			}

		}
		
		
		
		
		// 退辅料T+系统的选择模板并选择模板excel导出
				@RequestMapping("/outportSecMaterialExcel")
				@ResponseBody
				public String outportSecMaterialExcel(
						HttpServletRequest request,HttpServletResponse response,
						@RequestParam(value = "excelPath", required = false) MultipartFile excelPath){
				
					// 获取当天日期
					Date day = new Date();
					SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
					SimpleDateFormat df2 = new SimpleDateFormat("yyyy-MM");
					
					try {
						Workbook book1 = Workbook.getWorkbook(excelPath.getInputStream());// 获取流(会自动识别流中文件的实际地址，取得excel文件中的内容)
						String path = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
						String url=path+df1.format(day)+"退辅料单.xls";
						//String url="C:/Users/Administrator/Desktop/"+df1.format(day)+"退辅料单.xls";
						WritableWorkbook book= Workbook.createWorkbook(new File(url),book1);// 在原模板数据源的excel上追加数据
						WritableSheet sheet=book.getSheet(0);
						
				
						// 居中
						WritableCellFormat wf = new WritableCellFormat();
						wf.setAlignment(Alignment.CENTRE);
						
							
						// 判断原模板sheet1的名称是不是设备表，以确定导入的是不是材料出库单
						if (sheet.getName().equals("材料出库单") == false) {
							return "2";
						}
						
						for (int i = 1; i < sheet.getRows(); i++) {
							// 有效数据的行数
							Cell[] cell = sheet.getRow(i);
							// 判断有效数据的行中是否存在空
							if (cell.length == 0)
								continue;

							List<GetSecDetail> listsecoutMaterial=new ArrayList<GetSecDetail>();
						
							//退辅料的记录选择（除了外协的原材料出库）
							listsecoutMaterial = secMaterialService.selectexceloutputSecGetMaterial();
							
							if(listsecoutMaterial != null && listsecoutMaterial.size() > 0){
										// 5.将历史数据添加到单元格中 (先列后行)
										for (int j = 0; j < listsecoutMaterial.size(); j++) {
											//1.单据日期
											sheet.addCell((WritableCell) new Label(0, j + 2,df1.format(day), wf));
											
											//2.单据编号
											//（需要先查询数据库中该月中原材料的单据日期数量）
											Integer num=0;
											String num1=tJdbcService.selectRowMaterialNum();
											if(num1!=null&&"".equals(num1)){
											    num=Integer.valueOf(num1);
											}else{
												num=0;
											}
											
											//判断如果是1位数拼接为：当前日期到月份"yyyy-MM-+000+num"
											if(num>=0&&num<10){
												num=num+1+j;
												sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-000"+num, wf));
											}
											//判断如果2位数拼接为：当前日期到月份"yyyy-MM-+00+num"
											if(num>=10&&num<100){
												sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-00"+num, wf));
											}
											//判断如果3位数拼接为：当前日期到月份"yyyy-MM-+0+num"
											if(num>=100){
												sheet.addCell((WritableCell) new Label(1, j + 2,"MD-"+df2.format(day)+"-0"+num+j, wf));
											}
								
											//3.业务类型编码（领料默认为41，退料默认为42）
											sheet.addCell((WritableCell) new Label(2, j + 2,42+"", wf));
											//4.生产车间
											sheet.addCell((WritableCell) new Label(7, j + 2,listsecoutMaterial.get(j).getReshopName(), wf));
											//5.仓库编码
											sheet.addCell((WritableCell) new Label(11, j + 2,"03", wf));
											//6.材料编码
											sheet.addCell((WritableCell) new Label(15, j + 2,listsecoutMaterial.get(j).getSecMaterialNo(), wf));
											//7.计量单位
											sheet.addCell((WritableCell) new Label(18, j + 2,listsecoutMaterial.get(j).getUnit(), wf));
											//8.数量
											sheet.addCell((WritableCell) new Label(19, j + 2,listsecoutMaterial.get(j).getNum(), wf));
											//10.退辅料人
											sheet.addCell((WritableCell) new Label(23, j + 2,listsecoutMaterial.get(j).getReceiver(), wf));
											
										}
							}
								
						}
						// 6.写入excel并关闭
						book.write();
						book.close();
						return "1";
					} catch (Exception e) {
						return "0";

					}

				}
				
				
				//产成品入库单T+系统的选择模板并选择模板excel导出
				@RequestMapping("/outportFullProductExcel")
				@ResponseBody
				public String outportFullProductExcel(
						HttpServletRequest request,HttpServletResponse response,
						@RequestParam(value = "excelPath", required = false) MultipartFile excelPath)
						throws Exception {
					
					
					// 获取当天日期
					Date day = new Date();
					SimpleDateFormat df1 = new SimpleDateFormat("yyyy-MM-dd");
					SimpleDateFormat df2 = new SimpleDateFormat("yyyy-MM");
					
					try {
						Workbook book1 = Workbook.getWorkbook(excelPath.getInputStream());// 获取流(会自动识别流中文件的实际地址，取得excel文件中的内容)
						String path = FileSystemView.getFileSystemView().getHomeDirectory().getPath();
						String url=path+df1.format(day)+"产成品入库单.xls";
						//String url="C:/Users/Administrator/Desktop/"+df1.format(day)+"产成品入库单.xls";
						WritableWorkbook book= Workbook.createWorkbook(new File(url),book1);// 在原模板数据源的excel上追加数据
						WritableSheet sheet=book.getSheet(0);
						
						 
					
						// 居中
						WritableCellFormat wf = new WritableCellFormat();
						wf.setAlignment(Alignment.CENTRE);
						
							
						// 判断原模板sheet1的名称是不是设备表，以确定导入的是不是材料出库单
						if (sheet.getName().equals("产成品入库单") == false) {
							return "2";
						}
						
						for (int i = 1; i < sheet.getRows(); i++) {
							// 有效数据的行数
							Cell[] cell = sheet.getRow(i);
							// 判断有效数据的行中是否存在空
							if (cell.length == 0)
								continue;

							List<ProductRecord> listoutputmaterial = new ArrayList<ProductRecord>();
							
							//原材料的记录选择（除了外协的原材料出库）
							listoutputmaterial = inputOrOutputRecordService.selectexcelOutputFullRecord();
							
							if(listoutputmaterial != null && listoutputmaterial.size() > 0){
										// 5.将历史数据添加到单元格中 (先列后行)
										for (int j = 0; j < listoutputmaterial.size(); j++) {
											//1.单据日期
											sheet.addCell((WritableCell) new Label(0, j + 2,df1.format(day), wf));
											
											//2.单据编号
											//（需要先查询数据库中该月中原材料的单据日期数量）
											Integer num=0;
											String num1=tJdbcService.selectRowMaterialNum();
											if(num1!=null&&"".equals(num1)){
											    num=Integer.valueOf(num1);
											}else{
												num=0;
											}
											
											//判断如果是1位数拼接为：当前日期到月份"yyyy-MM-+000+num"
											if(num>=0&&num<10){
												num=num+1+j;
												sheet.addCell((WritableCell) new Label(1, j + 2,"MC-"+df2.format(day)+"-000"+num, wf));
											}
											//判断如果2位数拼接为：当前日期到月份"yyyy-MM-+00+num"
											if(num>=10&&num<100){
												sheet.addCell((WritableCell) new Label(1, j + 2,"MC-"+df2.format(day)+"-00"+num, wf));
											}
											//判断如果3位数拼接为：当前日期到月份"yyyy-MM-+0+num"
											if(num>=100){
												sheet.addCell((WritableCell) new Label(1, j + 2,"MC-"+df2.format(day)+"-0"+num+j, wf));
											}
								
											//3.业务类型编码（领料默认为41，退料默认为42）
											sheet.addCell((WritableCell) new Label(2, j + 2,"03", wf));
											//4.生产车间
											sheet.addCell((WritableCell) new Label(7, j + 2,listoutputmaterial.get(j).getShop1(), wf));
											//5.产成品入库人
											sheet.addCell((WritableCell) new Label(9, j + 2,listoutputmaterial.get(j).getProvider(), wf));
											//6.仓库编码
											sheet.addCell((WritableCell) new Label(11, j + 2,"02", wf));
											//7.产品编码
											sheet.addCell((WritableCell) new Label(15, j + 2,listoutputmaterial.get(j).getMaterialNo(), wf));
											//8.计量单位
											sheet.addCell((WritableCell) new Label(19, j + 2,listoutputmaterial.get(j).getUnit(), wf));
											//9.数量
											sheet.addCell((WritableCell) new Label(20, j + 2,listoutputmaterial.get(j).getProductNum(), wf));
										
											
										}
							}
								
						}
						// 6.写入excel并关闭
						book.write();
						book.close();
						return "1";
					} catch (Exception e) {
						return "0";

					}

				}
}