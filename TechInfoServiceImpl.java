package egovframework.home.tech.service.impl;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.ui.ModelMap;
import org.springframework.web.multipart.MultipartFile;

import com.fasterxml.jackson.databind.ObjectMapper;

import egovframework.com.cmm.ComDefaultCodeVO;
import egovframework.com.cmm.service.CmmnDetailCode;
import egovframework.com.cmm.service.EgovCmmUseService;
import egovframework.com.cmm.service.EgovFileMngService;
import egovframework.com.cmm.service.EgovFileMngUtil;
import egovframework.com.cmm.service.FileVO;
import egovframework.com.cop.bbs.service.impl.EgovArticleDAO2;
import egovframework.com.cop.ems.service.SndngMailVO;
import egovframework.com.cop.ems.service.impl.SndngMailRegistDAO;
import egovframework.com.sym.ccm.cca.service.CmmnCodeVO;
import egovframework.com.sym.ccm.cca.service.EgovCcmCmmnCodeManageService;
import egovframework.com.sym.log.wlg.service.WebLog;
import egovframework.com.sym.log.wlg.service.impl.WebLogDAO;
import egovframework.home.application.service.CnsVO;
import egovframework.home.search.SearchCommon;
import egovframework.home.search.vo.SearchAutoCompleteVO;
import egovframework.home.search.vo.SearchCommonVO;
import egovframework.home.send.service.SendVO;
import egovframework.home.tech.service.TechDashBoardVO;
import egovframework.home.tech.service.TechDealInfoVO;
import egovframework.home.tech.service.TechInfoService;
import egovframework.home.tech.service.TechInfoVO;
import egovframework.home.tech.service.TechPantentVO;
import egovframework.home.tech.service.TechProjectVO;
import egovframework.home.user.service.UserInfoService;
import egovframework.home.user.service.UserInfoVO;
import egovframework.rte.fdl.cmmn.EgovAbstractServiceImpl;
import egovframework.rte.fdl.idgnr.EgovIdGnrService;
import egovframework.rte.fdl.property.EgovPropertyService;
import egovframework.rte.psl.dataaccess.util.EgovMap;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.util.PDFTextStripper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

@Service("techInfoService")
public class TechInfoServiceImpl extends EgovAbstractServiceImpl implements TechInfoService  {

	@Resource(name = "techInfoDAO")
    private TechInfoDAO techInfoDAO;

	@Resource(name = "webLogDAO")
    private WebLogDAO webLogDAO;
	
	@Resource(name = "sndngMailRegistDAO")
    private SndngMailRegistDAO sndngMailRegistDAO;
	
	
	@Resource(name = "egovArticleDao2")
    private EgovArticleDAO2 egovArticleDao2;

    @Resource(name = "propertiesService")
    protected EgovPropertyService propertyService;
    
	@Resource(name = "techIdGnrService")
	private EgovIdGnrService techIdGnrService;
	
    
	@Resource(name = "userInfoService")
	private UserInfoService userInfoService;
    
	@Resource(name = "projectIdGnrService")
	private EgovIdGnrService projectIdGnrService;
	
	
	@Resource(name = "CmmnCodeManageService")
	private EgovCcmCmmnCodeManageService cmmnCodeManageService;

	@Resource(name = "EgovCmmUseService")
	private EgovCmmUseService cmmUseService;
	
	@Resource(name = "pentnetIdGnrService")
	private EgovIdGnrService pentnetIdGnrService;
	

    @Resource(name = "EgovFileMngUtil")
    private EgovFileMngUtil fileUtil;
    SearchCommon searchCommon = new SearchCommon();
    
	@Resource(name = "EgovFileMngService")
    private EgovFileMngService fileMngService;
	
	
	
	@Override
	public int insertPantentInfo(TechPantentVO techPantentVO) {
		
		return techInfoDAO.insertPantentInfo(techPantentVO);
	}

	@Override
	public int insertProjectInfo(TechProjectVO techProjectVO) {
		
		return techInfoDAO.insertProjectInfo(techProjectVO);
	}
	
	@Override
	public int selectPantentInfoCnt(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectPantentInfoCnt(techPantentVO);
	}
	
	@Override
	public List<?> selectPantentInfoList(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectPantentInfoList(techPantentVO);
	}
	
	@Override
	public TechPantentVO selectPantentInfo(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectPantentInfo(techPantentVO);
	}
	
	@Override
	public void insertTechInfo(TechInfoVO techInfoVO) {
		
		/**신청자에게 발송  **/
		 if(techInfoVO.getAnswer1().contains("Y")){     //이메일전송
			  SendVO sendVO = new  SendVO(); 
			  sendVO.setMailGubun("996");
			  sendVO.setTransRecipientNm(techInfoVO.getVenNm());
			  sendVO.setTransRecipient(techInfoVO.getVenEmail()+"@"+techInfoVO.getVenEmail2());
			  sendVO.setTransTitl("[해양수산 기술거래 플랫폼] 기술 등록 신청 안내");
			  sendVO.setTransCont4("해양수산 기술거래 플랫폼에서 안내드립니다.\n" +"기술정보 등록이 신청되었습니다.\n" +"등록 완료 시 알림 메시지가 발송됩니다.\n");	//메일내용
			  sendVO.setTransCont6(techInfoVO.getTechNm());	//기술명
			  sendVO.setTransCont10("https://ofris.kimst.re.kr/tech-trade");	//URL
			  egovArticleDao2.sendEMAIL(sendVO); 
          	
		  }
		  
		  //문자전송 
		 if(techInfoVO.getAnswer2().contains("Y")){  
			 String mmsContent = "[Web발신]\n" +"해양수산 기술거래 플랫폼에서 안내드립니다\n"+"기술정보 등록이 신청되었습니다.\n" +"등록 완료 시 알림 메시지가 발송됩니다.\n\r"; 
			 mmsContent += "기술명 : " + techInfoVO.getTechNm() + "\n\r"; 
			 mmsContent += "홈페이지 바로가기 : https://ofris.kimst.re.kr/tech-trade/\n\r\n\r";  
			 mmsContent += "감사합니다.";
			 SendVO sendVO = new SendVO();
			 sendVO.setTransRecipientNm(techInfoVO.getVenNm());
			 sendVO.setTransRecipient(techInfoVO.getVenHp1()+techInfoVO.getVenHp2()+techInfoVO.getVenHp3());
			 sendVO.setTransType(6);
			 sendVO.setTransTitl("해양수산과학기술진흥원"); 
			 sendVO.setTransCont4(mmsContent);
			 egovArticleDao2.sendMMS(sendVO); 
		 }
		
		 
		 /**관리자에게 발송  **/
		 String[] emails = {"yhhan1428@kimst.re.kr"};
		 SendVO sendVO = new  SendVO(); 
		 sendVO.setMailGubun("996");
		 sendVO.setTransRecipientNm("관리자");
		 sendVO.setTransTitl("[해양수산 기술거래 플랫폼] 기술 승인 요청 안내");
		 sendVO.setTransCont4("해양수산 기술거래 플랫폼에서 안내드립니다.\n" +"기술 등록 승인 요청이 접수되었습니다.\n");	//메일내용
		 sendVO.setTransCont6(techInfoVO.getTechNm());	//기술명
		 sendVO.setTransCont10("https://ofris.kimst.re.kr/tech-trade");
		 
		 for(int i=0;i<emails.length;i++){
			 sendVO.setTransRecipient(emails[i]);
			 egovArticleDao2.sendEMAIL(sendVO); 
		 }
		 
//		 20240430 주석처리_강동훈
		 /*String mmsContent = "[Web발신]\n" +"해양수산 기술거래 플랫폼에서 안내드립니다\n"+"기술 등록 승인 요청이 접수되었습니다.\n\r상세정보\n\r"; 
		 mmsContent += "기술명 : " + techInfoVO.getTechNm() + "\n"; 
		 mmsContent += "홈페이지 바로가기 : https://ofris.kimst.re.kr/tech-trade/\n\r";  
		 mmsContent += "감사합니다.";
		 SendVO sendVO2 = new SendVO();
		 sendVO2.setTransRecipientNm(techInfoVO.getVenNm());
		 sendVO2.setTransRecipient("01077089138");
		 sendVO2.setTransTitl("해양수산과학기술진흥원"); 
		 sendVO2.setTransType(6);
		 sendVO2.setTransCont4(mmsContent);*/
		 //egovArticleDao2.sendMMS(sendVO2); 
		 
		 
		techInfoDAO.insertTechInfo(techInfoVO);
	}
	
	@Override
	public List<?> selectTechInfoList(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectTechInfoList(techPantentVO);
	}
	public List<?> selectTechInfoListUser(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectTechInfoListUser(techPantentVO);
	}
	@Override
	public int selectTechInfoCnt(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectTechInfoCnt(techPantentVO);
	}
	public int selectTechInfoCntUser(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectTechInfoCntUser(techPantentVO);
	}
	@Override
	public TechInfoVO selectTechInfo(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectTechInfo(techInfoVO);
	}
	
	@Override
	public void updateTechCnt(TechInfoVO techInfoVO) {
		
		 techInfoDAO.updateTechCnt(techInfoVO);
	}
	
	@Override
	public List<?> selectAdminTechInfoList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectAdminTechInfoList(techInfoVO);
	}
	
	@Override
	public List<?> selectAdminApproveTechInfoList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectAdminApproveTechInfoList(techInfoVO);
	}
	
	@Override
	public List<TechInfoVO> selectAdminTechInfoExcelList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectAdminTechInfoExcelList(techInfoVO);
	}
	
	@Override
	public int selectAdminTechInfoCnt(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectAdminTechInfoCnt(techInfoVO);
	}
	
	@Override
	public int selectAdminApproveTechInfoCnt(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectAdminApproveTechInfoCnt(techInfoVO);
	}
	
	@Override
	public TechProjectVO selectTechProjectInfo(UserInfoVO userInfoVO) throws Exception {
		return techInfoDAO.selectTechProjectInfo(userInfoVO);
	}
	@Override
	public int updateApproveYn(TechInfoVO techInfoVO) { 
		String state="";
		String contents="";
		String contents2="";
		if(techInfoVO.getApproveYn().equals("Y")) {
			state="승인";
			contents="기술정보 등록이 완료되었습니다.";
			contents2="기술정보 등록이 완료되었습니다.";
		}else if(techInfoVO.getApproveYn().equals("R")) {
			state="반려";
			contents="기술정보 신청이 반려되었습니다.";
			contents2="기술정보 신청이 반려되었습니다.";
		}else if(techInfoVO.getApproveYn().equals("C")) {
			state="수정";
			contents="기술 수정 요청이 접수되었습니다.";
			contents2="기술 수정 요청이 접수되었습니다.";
		}else if(techInfoVO.getApproveYn().equals("E")) {
			state="서비스제외";
			contents="서비스제외 요청이 접수되었습니다.";
			contents2="서비스제외 요청이 접수되었습니다.";
		}
		
		if(techInfoVO.getAnswer1() == null) {
			techInfoVO.setAnswer1("");
		}
		if(techInfoVO.getAnswer2() == null) {
			techInfoVO.setAnswer2("");
		}
		
		if(techInfoVO.getApproveYn() !="E") {
		 if(techInfoVO.getAnswer1().contains("Y")){   //이메일전송
			  SendVO sendVO = new  SendVO(); 
			  sendVO.setMailGubun("996");
			  sendVO.setTransRecipientNm(techInfoVO.getVenNm());
			  sendVO.setTransRecipient(techInfoVO.getVenEmail()+"@"+techInfoVO.getVenEmail2());
			  sendVO.setTransTitl("[해양수산 기술거래 플랫폼] 기술 "+state+" 안내");
			  sendVO.setTransCont4("해양수산 기술거래 플랫폼에서 안내드립니다.\n"+contents+"\n\r");	//메일내용
			  sendVO.setTransCont6(techInfoVO.getTechNm());	//기술명
			  sendVO.setTransCont10("https://ofris.kimst.re.kr/tech-trade");	//URL
			  egovArticleDao2.sendEMAIL(sendVO); 
         	
		  }
		  
		 
		  //문자전송 
		 if(techInfoVO.getAnswer2().contains("Y")){
			 String mmsContent = "[Web발신]\n" +"해양수산 기술거래 플랫폼에서 안내드립니다\n"+contents2+"\n\r"; 
			 mmsContent += "기술명 : " + techInfoVO.getTechNm() + "\n\r"; 
			 mmsContent += "홈페이지 바로가기 : https://ofris.kimst.re.kr/tech-trade/\n\r";  
			 mmsContent += "감사합니다.";
			 SendVO sendVO = new SendVO();
			 sendVO.setTransRecipientNm(techInfoVO.getVenNm());
			 sendVO.setTransRecipient(techInfoVO.getVenHp1()+techInfoVO.getVenHp2()+techInfoVO.getVenHp3());
			 sendVO.setTransTitl("해양수산과학기술진흥원"); 
			 sendVO.setTransCont4(mmsContent);
			 sendVO.setTransType(6);
			 egovArticleDao2.sendMMS(sendVO); 
		 }
		}else {
			 if(techInfoVO.getAnswer1().contains("Y")){   //이메일전송
				  SendVO sendVO = new  SendVO(); 
				  sendVO.setMailGubun("996");
				  sendVO.setTransRecipientNm(techInfoVO.getVenNm());
				  sendVO.setTransRecipient(techInfoVO.getVenEmail()+"@"+techInfoVO.getVenEmail2());
				  sendVO.setTransTitl("[해양수산 기술거래 플랫폼] 기술"+state+ " 안내");
				  sendVO.setTransCont4("해양수산 기술거래 플랫폼에서 안내드립니다.\n"+contents+"\n\r");	//메일내용
				  sendVO.setTransCont6(techInfoVO.getTechNm());	//기술명
				  sendVO.setTransCont10("https://ofris.kimst.re.kr/tech-trade");	//URL
				  egovArticleDao2.sendEMAIL(sendVO); 
	         	
			  }
			  
			  //문자전송 
			 if(techInfoVO.getAnswer2().contains("Y")){
				 String mmsContent = "[Web발신]\n" +"해양수산 기술거래 플랫폼에서 안내드립니다.\n\r"+contents2+"\n\r"; 
				 mmsContent += "상세정보\n\r기술명 : " + techInfoVO.getTechNm() + "\n\r"; 
				 mmsContent += "홈페이지 바로가기 : https://ofris.kimst.re.kr/tech-trade/\n";  
				 mmsContent += "감사합니다.";
				 SendVO sendVO = new SendVO();
				 sendVO.setTransRecipientNm(techInfoVO.getVenNm());
				 sendVO.setTransRecipient(techInfoVO.getVenHp1()+techInfoVO.getVenHp2()+techInfoVO.getVenHp3());
				 sendVO.setTransTitl("해양수산과학기술진흥원"); 
				 sendVO.setTransCont4(mmsContent);
				 sendVO.setTransType(6);
				 egovArticleDao2.sendMMS(sendVO); 
			 }
		}
		
		/**관리자에게 발송  **/
		 String[] emails = {"yhhan1428@kimst.re.kr"};
		 SendVO sendVO = new  SendVO(); 
		 sendVO.setMailGubun("996");
		 sendVO.setTransRecipientNm(techInfoVO.getVenNm());
		 sendVO.setTransTitl("[해양수산 기술거래 플랫폼] 기술"+state+ " 안내");
		 sendVO.setTransCont4("해양수산 기술거래 플랫폼에서 안내드립니다.\n"+contents+"\n\r");	//메일내용
		 sendVO.setTransCont6(techInfoVO.getTechNm());	//기술명
		 sendVO.setTransCont10("https://ofris.kimst.re.kr/tech-trade");
		 
		 for(int i=0;i<emails.length;i++){
			 sendVO.setTransRecipient(emails[i]);
			 egovArticleDao2.sendEMAIL(sendVO); 
		 }
		 /* 20240430 주석처리_강동훈
		 String mmsContent = "[Web발신]\n" +"해양수산 기술거래 플랫폼에서 안내드립니다.\n\r"+contents2+"\n\r"; 
		 mmsContent += "상세정보\n\r기술명 : " + techInfoVO.getTechNm() + "\n\r"; 
		 mmsContent += "홈페이지 바로가기 : https://ofris.kimst.re.kr/tech-trade/\n\r";  
		 mmsContent += "감사합니다.";
		 SendVO sendVO2 = new SendVO();
		 sendVO2.setTransRecipientNm(techInfoVO.getVenNm());
		 sendVO2.setTransRecipient("01077089138");
		 sendVO2.setTransTitl("해양수산과학기술진흥원"); 
		 sendVO2.setTransCont4(mmsContent);
		 sendVO2.setTransType(6);*/
		 //egovArticleDao2.sendMMS(sendVO2);
		 
		 int success = techInfoDAO.updateApproveYn(techInfoVO);
		 
		 
		 
		return  success;
	}
	
	@Override
	public int selectMainCnt(WebLog webLog) throws Exception {
		
		return  webLogDAO.selectWebLogInfCnt(webLog);
	}
	
	@Override
	public int selectDealCateCdCnt(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectDealCateCdCnt(techInfoVO);
	}
	
	public int selectDealCateCdCntUser(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectDealCateCdCntUser(techInfoVO);
	}
	
	@Override
	public List<?> selectCateStLcdList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectCateStLcdList(techInfoVO);
	}
	public List<?> selectCateStLcdListUser(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectCateStLcdListUser(techInfoVO);
	}
	
	@Override
	public int selectCateStLcdCnt(TechInfoVO techInfoVO) throws Exception {
		
		return  techInfoDAO.selectCateStLcdCnt(techInfoVO);
	}
	
	@Override
	public List<?> selectCateOcLcdList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectCateOcLcdList(techInfoVO);
	}
	
	public List<?> selectCateOcLcdListUser(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectCateOcLcdListUser(techInfoVO);
	}
	
	@Override
	public int selectCateOcLcdCnt(TechInfoVO techInfoVO) throws Exception {
		
		return  techInfoDAO.selectCateOcLcdCnt(techInfoVO);
	}
	
	@Override
	public List<?> selectPantentRgnNmList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectPantentRgnNmList(techInfoVO);
	}
	
	@Override
	public List<?> selectAgencyList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectAgencyList(techInfoVO);
	}
	public List<?> selectAgencyListUser(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectAgencyListUser(techInfoVO);
	}
	
	@Override
	public int selectAgencyCnt(TechInfoVO techInfoVO) throws Exception {
		
		return  techInfoDAO.selectAgencyCnt(techInfoVO);
	}
	public int selectAgencyCntUser(TechInfoVO techInfoVO) throws Exception {
		
		return  techInfoDAO.selectAgencyCntUser(techInfoVO);
	}
	@Override
	public List<?> selectPantentYearList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectPantentYearList(techInfoVO);
	}
	public List<?> selectPantentYearListUser(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectPantentYearListUser(techInfoVO);
	}
	
	@Override
	public int selectPantentYearCnt(TechInfoVO techInfoVO) throws Exception {
		
		return  techInfoDAO.selectPantentYearCnt(techInfoVO);
	}
	
	@Override
	public List<?> selectAllPantentInfoList(TechPantentVO techPantentVO) throws Exception {
		return techInfoDAO.selectAllPantentInfoList(techPantentVO);
	}
	
	@Override
	public void updatePantentImg(TechPantentVO techPantentVO) {
		
		 techInfoDAO.updatePantentImg(techPantentVO);
	}
	
	@Override
	public int updateTechInfo(TechInfoVO techInfoVO) throws Exception {
		
		/**신청자에게 발송  **/
		if(techInfoVO.getAnswer1() != null) {
			 if(techInfoVO.getAnswer1().contains("Y")){     //이메일전송
				  SendVO sendVO = new  SendVO(); 
				  sendVO.setMailGubun("996");
				  sendVO.setTransRecipientNm(techInfoVO.getVenNm());
				  sendVO.setTransRecipient(techInfoVO.getVenEmail()+"@"+techInfoVO.getVenEmail2());
				  sendVO.setTransTitl("[해양수산 기술거래 플랫폼] 기술 수정 요청 안내");
				  sendVO.setTransCont4("해양수산 기술거래 플랫폼에서 안내드립니다.\n" +"기술 수정 요청이 접수되었습니다.\n");	//메일내용
				  sendVO.setTransCont6(techInfoVO.getTechNm());	//기술명
				  sendVO.setTransCont10("https://ofris.kimst.re.kr/tech-trade");	//URL
				  egovArticleDao2.sendEMAIL(sendVO); 
	         	
			  }
		}
		  
		  //문자전송 
		if(techInfoVO.getAnswer2() != null) {
			 if(techInfoVO.getAnswer2().contains("Y")){  
				 String mmsContent = "[Web발신]\n" +"해양수산 기술거래 플랫폼에서 안내드립니다.\n"+"기술 수정 요청이 접수되었습니다.\n"; 
				 mmsContent += "기술명 : " + techInfoVO.getTechNm() + "\n\r"; 
				 mmsContent += "홈페이지 바로가기 : https://ofris.kimst.re.kr/tech-trade/\n\r\n\r";  
				 mmsContent += "감사합니다.";
				 SendVO sendVO = new SendVO();
				 sendVO.setTransRecipientNm(techInfoVO.getVenNm());
				 sendVO.setTransRecipient(techInfoVO.getVenHp1()+techInfoVO.getVenHp2()+techInfoVO.getVenHp3());
				 sendVO.setTransTitl("해양수산과학기술진흥원"); 
				 sendVO.setTransCont4(mmsContent);
				 sendVO.setTransType(6);
				 egovArticleDao2.sendMMS(sendVO); 
			 }
		}
		
		/**관리자에게 발송  **/
		 String[] emails = {"yhhan1428@kimst.re.kr"};
		 SendVO sendVO = new  SendVO(); 
		 sendVO.setMailGubun("996");
		 sendVO.setTransRecipientNm(techInfoVO.getVenNm());
		 sendVO.setTransTitl("[해양수산 기술거래 플랫폼] 기술 수정 요청 안내");
		 sendVO.setTransCont4("해양수산 기술거래 플랫폼에서 안내드립니다.\n" +"기술 수정 요청이 접수되었습니다.\n");	//메일내용
		 sendVO.setTransCont6(techInfoVO.getTechNm());	//기술명
		 sendVO.setTransCont10("https://ofris.kimst.re.kr/tech-trade");
		 
		 for(int i=0;i<emails.length;i++){
			 sendVO.setTransRecipient(emails[i]);
			 egovArticleDao2.sendEMAIL(sendVO); 
		 }
		  /* 20240430 주석처리_강동훈
		 String mmsContent = "[Web발신]\n" +"해양수산 기술거래 플랫폼에서 안내드립니다\n"+"기술 수정 요청이 접수되었습니다.\n상세정보\n\r"; 
		 mmsContent += "기술명 : " + techInfoVO.getTechNm() + "\n\r"; 
		 mmsContent += "홈페이지 바로가기 : https://ofris.kimst.re.kr/tech-trade/\n\r\n\r";  
		 mmsContent += "감사합니다.";
		 SendVO sendVO2 = new SendVO();
		 sendVO2.setTransRecipientNm(techInfoVO.getVenNm());
		 sendVO2.setTransRecipient("01077089138");
		 sendVO2.setTransTitl("해양수산과학기술진흥원"); 
		 sendVO2.setTransCont4(mmsContent);
		 sendVO2.setTransType(6);*/
		 //egovArticleDao2.sendMMS(sendVO2); 
		
		return techInfoDAO.updateTechInfo(techInfoVO);
	}

	@Override
	public int deleteTechInfo(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.deleteTechInfo(techInfoVO);
	}
	
	@Override
	public void updateTechProjectInfo(TechProjectVO techProjectVO) throws Exception {
		 techInfoDAO.updateTechProjectInfo(techProjectVO);
	}
	
	@Override
	public int selectTechExhiditionCnt(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectTechExhiditionCnt(techInfoVO);
	}
	
	@Override
	public List<?> selectTechExhiditionList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectTechExhiditionList(techInfoVO);
	}
	
	@Override
	public List<?> movingFileDBList(FileVO fileVO) throws Exception {
		
		List<?> list = techInfoDAO.movingFileDBList(fileVO);
		return list;
	}

	@Override
	public void movingFileInsert(FileVO fileVO) throws Exception {
		techInfoDAO.movingFileInsert(fileVO);
		
	}

	@Override
	public void movingFileDetailInsert(FileVO fileVO) throws Exception {
		techInfoDAO.movingFileDetailInsert(fileVO);
		
	}

	@Override
	public void updateTechSMK(TechInfoVO techInfoVO) {
		techInfoDAO.updateTechSMK(techInfoVO);
		
	}

	@Override
	public int checkExhidition(TechInfoVO techInfoVO) throws Exception {
		int cnt = techInfoDAO.checkExhidition(techInfoVO);
		return cnt;
	}

	@Override
	public void updateExhidition(TechInfoVO techInfoVO) {
		techInfoDAO.updateExhidition(techInfoVO);
		
	}

	@Override
	public void updatePreExhidition(TechInfoVO techInfoVO) {
		techInfoDAO.updatePreExhidition(techInfoVO);
		
	}

	@Override
	public void deletePatent(TechPantentVO techPatentVO) {
		techInfoDAO.deletePatent(techPatentVO);
		
	}
	
	@Override
	public void updatePatent(TechPantentVO techPatentVO) {
		techInfoDAO.updatePatent(techPatentVO);
		
	}

	@Override
	public int selectTechVideoCnt(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectTechVideoCnt(techInfoVO);
	}

	@Override
	public List<?> selectTechVideoList(TechInfoVO techInfoVO) throws Exception {
		List<?> list = techInfoDAO.selectTechVideoList(techInfoVO);
		return list;
	}

	@Override
	public void updateTechVideo(TechInfoVO techInfoVO) {
		techInfoDAO.updateTechVideo(techInfoVO);
		
	}

	@Override
	public void updatePreTechVideo(TechInfoVO techInfoVO) {
		techInfoDAO.updatePreTechVideo(techInfoVO);
		
	}

	@Override
	public void updateExhiditionViewCnt(TechInfoVO techInfoVO) {
		techInfoDAO.updateExhiditionViewCnt(techInfoVO);
		
	}
	
	@Override
	public List<?> selectMainTechList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectMainTechList(techInfoVO);
	}
	
	@Override
	public List<?> selectMainTechExhiditionList(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectMainTechExhiditionList(techInfoVO);
	}
	
	@Override
	public int myCounselingCnt(CnsVO cnsVO) throws Exception {
		return techInfoDAO.myCounselingCnt(cnsVO);
	}
	

	public Map selectDealCateFacetMap(TechInfoVO techInfoVO) throws Exception{
		return techInfoDAO.selectDealCateFacetMap(techInfoVO);
	}

	@Override
	public List<?> selectApprovalTop5List(TechDashBoardVO TDVO) throws Exception {
		List<?> list = techInfoDAO.selectApprovalTop5List(TDVO);
		return list;
	}
	
	public List<TechInfoVO> selectTechInfoExcelListUser(TechInfoVO techInfoVO) throws Exception{
		return techInfoDAO.selectTechInfoExcelListUser(techInfoVO);
	}
	public void insertTechDealInfo(TechDealInfoVO techDealInfoVO)throws Exception{
		techInfoDAO.insertTechDealInfo(techDealInfoVO);
	}
	
	public int updateTechDealStCd(TechInfoVO techInfoVO)throws Exception{
		return techInfoDAO.updateTechDealStCd(techInfoVO);
	}

	public int selectTechExhiditionCntUser(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectTechExhiditionCntUser(techInfoVO);
	}

	public List<?> selectTechExhiditionListUser(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectTechExhiditionListUser(techInfoVO);
	}
	
	public int selectTechVideoCntUser(TechInfoVO techInfoVO) throws Exception{
		return techInfoDAO.selectTechVideoCntUser(techInfoVO);
	}
	
	public List<?> selectTechVideoListUser(TechInfoVO techInfoVO) throws Exception{
		return techInfoDAO.selectTechVideoListUser(techInfoVO);
	}
	
	public List<EgovMap> selectPrevNextList(TechInfoVO techInfoVO) throws Exception{
		return techInfoDAO.selectPrevNextList(techInfoVO);
	}

	@Override
	public int approvalTop5Cnt(TechDashBoardVO TDVO) throws Exception {
		
		return techInfoDAO.approvalTop5Cnt(TDVO);
	}

	@Override
	public List<?> selectAllApprovalList(TechDashBoardVO TDVO) throws Exception {
		List<?> list = techInfoDAO.selectAllApprovalList(TDVO);
		
		return list;
	}

	@Override
	public int selectAllApprovalCnt(TechDashBoardVO TDVO) throws Exception {
		
		return techInfoDAO.selectAllApprovalCnt(TDVO);
	}

	public Map selectPrevNextVideoList(TechInfoVO techInfoVO) throws Exception{
		return techInfoDAO.selectPrevNextVideoList(techInfoVO);
	}
	
	public int checkPantentInfo(TechPantentVO techPatentVO) throws Exception {
		return techInfoDAO.checkPantentInfo(techPatentVO);
	}
	
	@Override
	public List<?> selectFileDetailList(SndngMailVO tndngMailVO) throws Exception {
		
		return sndngMailRegistDAO.selectAtchmnFileList(tndngMailVO);
	}
	
	@Override
	public int deleteTech(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.deleteTech(techInfoVO);
	}
	
	@Override
	public int deleteOneTech(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.deleteOneTech(techInfoVO);
	}
	
	@Override
	public List<?> selectMyTechInfoList(TechInfoVO techInfoVO) throws Exception {
		
		return techInfoDAO.selectMyTechInfoList(techInfoVO);
	}
	
	public int selectMyTechInfoCnt(TechInfoVO techInfoVO) throws Exception {
		return techInfoDAO.selectMyTechInfoCnt(techInfoVO);
	}
	
	public TechDealInfoVO selectTechDealInfo(TechDealInfoVO techDealInfoVO) throws Exception{
		return techInfoDAO.selectTechDealInfo(techDealInfoVO);
	}
	
	public List<?> selectTechDealList(CnsVO cnsVO) throws Exception{
		
		return techInfoDAO.selectTechDealList(cnsVO);
	}
	
	public int selectTechDealCnt(CnsVO cnsVO)throws Exception{
		return techInfoDAO.selectTechDealCnt(cnsVO);
	}
	
	
	@Override
	public void insertTechVideo(TechInfoVO techInfoVO) {
		techInfoDAO.insertTechVideo(techInfoVO);
		
	}
	
	public Map selectTechVideoInfo(TechInfoVO techInfoVO) throws Exception{
		return techInfoDAO.selectTechVideoInfo(techInfoVO);
	}
	
	public void techVideoReadCnt(TechInfoVO techInfoVO) throws Exception{
		techInfoDAO.techVideoReadCnt(techInfoVO);
	}
	
	public int deleteTechVideo(TechInfoVO techInfoVO)throws Exception{
		return techInfoDAO.deleteTechVideo(techInfoVO);
	}

	@Override
	public Map selectMinMaxPatentYear() throws Exception {
		return techInfoDAO.selectMinMaxPatentYear();
	}
	
	@Override
	public List<?> selectProjectInfoNew(TechProjectVO TechProjectVO) throws Exception{		
		return techInfoDAO.selectProjectInfoNew(TechProjectVO);
	}
	
	public int selectProjectInfoNewCnt(TechProjectVO TechProjectVO) throws Exception{		
		return techInfoDAO.selectProjectInfoNewCnt(TechProjectVO);
	}
	
	
	@Override
    public Map<String, Object> processExcelData(MultipartHttpServletRequest multiRequest, TechInfoVO techInfoVO, UserInfoVO userInfoVO, TechProjectVO techProjectVO, ModelMap model) throws Exception {
      

	    Map<String, Object> result = new HashMap<>();
	    List<String> errors = new ArrayList<>();
	    List<TechInfoVO> techInfoList = new ArrayList<>();
	    List<UserInfoVO> userInfoList = new ArrayList<>();
	    List<TechPantentVO> techPantentInfoList = new ArrayList<>();
	    List<TechProjectVO> techProjectInfoList = new ArrayList<>();
	    int successCount = 0;
	    int failCount = 0;
	    String smkFileId = "";
    	String mainImg = "";
		  String attFile = "";
	  	 List<FileVO> smkResult = null;
	  	 List<FileVO> smkFilesResult = null;
	  	 List<FileVO> imgFilesResult = null;
	  	 List<FileVO> atchFilesResult = null;
	    String techCd = techInfoVO.getTechCd();	


	    try {
	        List<MultipartFile> files = multiRequest.getFiles("file_excel");
          	final List<MultipartFile> smkFile = multiRequest.getFiles("file_smk"); //smkFile
        	final List<MultipartFile> file2 = multiRequest.getFiles("file_img"); // 대표 이미지
        	final List<MultipartFile> file3 = multiRequest.getFiles("file_multi"); // 첨부파일
        	
	        if (files.isEmpty() || files.get(0).isEmpty()) {
	            errors.add("엑셀 파일을 선택해주세요.");
	        } else {
	            MultipartFile file = files.get(0);
	            try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
	                Sheet sheet = workbook.getSheetAt(0);
	                Iterator<Row> rowIterator = sheet.iterator();

	                int rowNum = 0;
	                while (rowIterator.hasNext()) {
	                	int techSn = techIdGnrService.getNextIntegerId();
	                    rowNum++;
	                    Row row = rowIterator.next();

	                    if (rowNum == 1) continue; // 헤더 행 건너뛰기

	                    List<String> rowValues = new ArrayList<>();
	                    int lastColumnIndexWithData = -1; // 데이터가 있는 마지막 열 인덱스 초기화

	                    // 모든 셀을 순회하며 데이터가 있는 마지막 열 인덱스 찾기
	                    for (int columnIndex = 0; columnIndex <= row.getLastCellNum(); columnIndex++) {
	                        Cell cell = row.getCell(columnIndex);
	                        if (cell != null && !cellToString(cell).isEmpty()) {
	                            lastColumnIndexWithData = columnIndex;
	                        }
	                    }

	                    // 찾은 마지막 열 인덱스까지 데이터를 rowValues에 추가
	                    for (int columnIndex = 0; columnIndex <= lastColumnIndexWithData; columnIndex++) {
	                        Cell cell = row.getCell(columnIndex);
	                        String cellValue = (cell != null) ? cellToString(cell) : "";
	                        rowValues.add(cellValue);
	                    }
	                    
	                    boolean rowSuccess = true;
	                    // 유효성 검사 시 lastColumnIndexWithData + 1 사용
	                    smkFileId = "";
	                    mainImg = "";
	                    attFile = "";
	            	  	  smkResult = null;
	            	  	  smkFilesResult = null;
	            	  	 imgFilesResult = null;
	            	  	  atchFilesResult = null;
	                    for (int i = 0; i < lastColumnIndexWithData + 1; i++) {
		                    	//거래자정보 공통필수입력정보 확인
	                    	if (rowValues.get(0).isEmpty() && i==0) {
	                    		 errors.add(rowNum + "행 " + (i + 1) + "열: 거래자 구분값이 비어있습니다.");
		                         rowSuccess = false;
	                    	}	                    	
	                    	if((rowValues.get(12).isEmpty() && i==12) || (rowValues.get(13).isEmpty() && i==12)) {
	                    		errors.add(rowNum + "행 : 이메일 입력값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if((rowValues.get(14).isEmpty() && i==14) || (rowValues.get(15).isEmpty() && i==15)|| (rowValues.get(16).isEmpty() && i==16)) {
	                    		errors.add(rowNum + "행 : 전화번호 입력값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if((rowValues.get(17).isEmpty() && i==17) || (rowValues.get(15).isEmpty() && i==18)|| (rowValues.get(16).isEmpty() && i==19)) {
	                    		errors.add(rowNum + "행 : 휴대전화번호 입력값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if((rowValues.get(20).isEmpty() && i==20) || (rowValues.get(21).isEmpty() && i==21)|| (rowValues.get(22).isEmpty() && i==22)) {
	                    		errors.add(rowNum + "행 : 주소 입력값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(23).isEmpty()&& i==23) {
	                    		errors.add(rowNum + "행 : 답변알림 입력값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(24).isEmpty() && i==24) {
	                    		errors.add(rowNum + "행 : 답변알림 입력값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	
	                    	if(!rowValues.get(23).isEmpty() && !rowValues.get(23).equals("Y") && !rowValues.get(23).equals("N")&&i==23){
	                    		 errors.add(rowNum + "행 " + (i + 1) + "열: 이메일 답변 코드를 확인해주세요.");
	                    	}
	              
	                    	if(!rowValues.get(24).isEmpty() && !rowValues.get(24).equals("Y") && !rowValues.get(24).equals("N")&&i==24){
	                    		 errors.add(rowNum + "행 " + (i + 1) + "열: 휴대전화 답변 코드를 확인해주세요.");
	                    	}
	                    	if (rowValues.get(9).isEmpty()&&i==9) {
	                    		 errors.add(rowNum + "행 " + (i + 1) + "열: 이름(담당자) 입력값이 비어있습니다.");
		                         rowSuccess = false;
	                    	}	   
	                    	
	                    	if ((rowValues.get(0).equals("01")&&i==0) || (rowValues.get(0).equals("02")&&i==1)) {
	                    		if(rowValues.get(1).isEmpty() && i==1) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 기관명/기업명이 입력값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(10).isEmpty()&& i==10) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 부서명 입력값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(11).isEmpty() && i==11) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 직책 입력값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    	}
	                    	if(rowValues.get(0).equals("01")) {
	                    		//기업구분(창업기업01, 소기업02, 중기업03, 중견기업04, 대기업05)
	                    		if(rowValues.get(2).isEmpty()&&i==2) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 기업 구분값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(!rowValues.get(2).isEmpty()&&!rowValues.get(2).equals("01")&&!rowValues.get(2).equals("02")
	                    		&&!rowValues.get(2).equals("03")&&!rowValues.get(2).equals("04")&&!rowValues.get(2).equals("05")&&i==2) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 기업구분 코드를 확인해주세요.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(3).isEmpty()&&i==3) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 이름(대표자)값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(4).isEmpty()&&i==4) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 사업자등록번호값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(7).isEmpty()&&i==7) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 기업분류(대)값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(8).isEmpty()&&i==8) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 기업분류(중)값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
                    			CmmnCodeVO vo = new CmmnCodeVO();
                    			vo.setSearchCondition("1");
                    			vo.setSearchKeyword(rowValues.get(7));
                    			vo.setFirstIndex(0);
                    			vo.setRecordCountPerPage(20);
//                    			vo.setCodeId(rowValues.get(7));
                    	        List<?> ocCode = cmmnCodeManageService.selectCmmnCodeList(vo);
//                    	        해양산업코드	COM105	해양자원 개발 및 건설업
//                    	        해양산업코드	COM106	해운항만업
//                    	        해양산업코드	COM107	선박 및 해양플랜트 건조수리업
//                    	        해양산업코드	COM108	수산물 생산업
//                    	        해양산업코드	COM110	수산물 유통업
//                    	        해양산업코드	COM111	해양수산 레저관광업
//                    	        해양산업코드	COM112	해양수산 기자재 제조업
//                    	        해양산업코드	COM113	해양수산 관련 서비스업
//                    	        해양산업코드	COM109	수산물 가공업
	                    		if(!rowValues.get(7).isEmpty()&&i==7) {

	                    	        if(ocCode == null || ocCode.isEmpty()) {
	                    	        	errors.add(rowNum + "행 " + (i + 1) + "열: 기업분류(대)코드를 확인해주세요.");
				                        rowSuccess = false;	                    	        	
	                    	        }
                    	        }
//	                    		COM105	COM105_03	항만 및 해상 교량 건설업
//	                    		COM105	COM105_04	해양 수산플랜트 및 구조물 공사업
//	                    		COM105	COM105_01	해양자원 생산, 공급 및 개발업
//	                    		COM106	COM106_01	해운업
//	                    		COM106	COM106_02	항만업
//	                    		COM107	COM107_01	선박 건조 및 수리업
//	                    		COM107	COM107_02	해양 플랜트, 구조물 건조 및 수리업
//	                    		COM108	COM108_01	어로어업
//	                    		COM108	COM108_02	양식어업
//	                    		COM108	COM108_03	어업 관련 서비스업
//	                    		COM108	COM108_04	소금 채취업
//	                    		COM110	COM110_01	수산물 중개 및 도소매업
//	                    		COM110	COM110_02	수산물 운송 및 보관업
//	                    		COM111	COM111_02	수산레저관광업
//	                    		COM113	COM113_04	해양수산인력 고용 알선 및 공급업
                    	        if(ocCode!=null&&!ocCode.isEmpty()&&!rowValues.get(8).isEmpty()&&!rowValues.get(7).isEmpty()&&i==8) {
                    	        	ComDefaultCodeVO defaultVO = new ComDefaultCodeVO();
                    	        	defaultVO.setCodeId(rowValues.get(7));
                    	        	List<CmmnDetailCode> codeDetailList = cmmUseService.selectCmmCodeDetail(defaultVO);
                    	        	boolean codeFound = false;
                                    for (Object item : codeDetailList) {
                                        if (item instanceof CmmnDetailCode) { // ComDefaultCodeVO 타입인지 확인
                                        	CmmnDetailCode codeVO = (CmmnDetailCode) item; // ComDefaultCodeVO로 캐스팅
                                            if (rowValues.get(8).equals(codeVO.getCode())) { // CODE 값을 비교
                                                codeFound = true;
//                                                break;
                                            }
                                        }
                                    }
                                    if (!codeFound) { // 세부 코드 목록에 rowValues.get(8)이 없는 경우
                                        errors.add(rowNum + "행 " + (i + 1) + "열: 기업분류(중) 코드를 확인해주세요.");
                                        rowSuccess = false;
                                    }		
                    	        }
	                    	       
	                    		}
	                    	//거래자정보 체크 마지막
	                    	//기술정보 체크 시작
	                    	//필수값(연구자 과제 유형, 과제명, 과제번호, 표준산업-대분류, 표준산업-중분류, 해양수산업-대분류, 해양수산업-중분류
	                    	//    기술명, 연구자명, 연구기관, 적용분야, 키워드, 요약, 기술 완성도, 거래유형, 기술료 조건, 거래상태정보, 기술의 우수성
	                    	if(rowValues.get(25).isEmpty()&&i==25) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 연구자 과제 유형값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(!rowValues.get(25).isEmpty()&&!rowValues.get(25).equals("01")&&!rowValues.get(25).equals("02")
	                    			&&!rowValues.get(25).equals("03")&&!rowValues.get(25).equals("04")&&i==25) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 연구자 과제유형 코드를 확인하세요.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(28).isEmpty()&&!rowValues.get(25).equals("04")&&!rowValues.get(25).equals("02")&&i==28) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 부처명값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(rowValues.get(29).isEmpty()&&!rowValues.get(25).equals("04")&&i==29) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 과제명값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(rowValues.get(30).isEmpty()&&!rowValues.get(25).equals("04")&&rowValues.get(25).equals("01")&&i==30) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 과제번호값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(rowValues.get(31).isEmpty()&&!rowValues.get(25).equals("04")&&rowValues.get(25).equals("03")&&i==31) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 수행연구기관명값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(rowValues.get(32).isEmpty()&&i==32) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 표준산업-대분류값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	CmmnCodeVO vo = new CmmnCodeVO();
                			vo.setSearchCondition("1");
                			vo.setSearchKeyword(rowValues.get(32));
                			vo.setFirstIndex(0);
                			vo.setRecordCountPerPage(20);
                			vo.setCodeId(rowValues.get(7));
//                			표준산업코드	COM117	전기, 가스, 증기 및 수도사업
//                			표준산업코드	COM118	하수, 폐기물 처리, 원료재생 및 환경복원업
//                			표준산업코드	COM119	건설업
//                			표준산업코드	COM120	도매 및 소매업
//                			표준산업코드	COM121	운수업
//                			표준산업코드	COM122	숙박 및 음식점업
//                			표준산업코드	COM123	출판, 영상, 방송통신 및 정보서비스업
//                			표준산업코드	COM124	금융 및 보험업
//                			표준산업코드	COM125	부동산업 및 임대업
//                			표준산업코드	COM126	전문, 과학 및 기술 서비스업
//                			표준산업코드	COM127	사업시설관리 및 사업지원 서비스업
//                			표준산업코드	COM128	공공행정, 국방 및 사회보장 행정
//                			표준산업코드	COM129	교육 서비스업
//                			표준산업코드	COM130	보건업 및 사회복지 서비스업
//                			표준산업코드	COM131	예술, 스포츠 및 여가관련 서비스업
//                			표준산업코드	COM132	협회 및 단체, 수리 및 기타 개인 서비스업
//                			표준산업코드	COM133	가구내 고용활동 및 달리 분류되지 않은 자가소비 생산활동
//                			표준산업코드	COM134	국제 및 외국기관
                	        List<?> ocCode = cmmnCodeManageService.selectCmmnCodeList(vo);
	                    	if(!rowValues.get(32).isEmpty()&&i==32) {
                    		
                    	        if(ocCode == null || ocCode.isEmpty()) {
                    	        	errors.add(rowNum + "행 " + (i + 1) + "열: 표준산업 대분류 코드를 확인해주세요.");
			                        rowSuccess = false;	                    	        	
                    	        }
                    	       
                    		}	
//	                    	COM117	COM117_01	전기, 가스, 증기 및 공기조절 공급업
//	                    	COM117	COM117_02	수도사업
//	                    	COM118	COM118_01	하수, 폐수 및 분뇨 처리업
//	                    	COM118	COM118_02	폐기물 수집운반, 처리 및 원료재생업
//	                    	COM124	COM124_02	보험 및 연금업
//	                    	COM124	COM124_03	금융 및 보험 관련 서비스업
//	                    	COM125	COM125_01	부동산업
//	                    	COM126	COM126_01	연구개발업
//	                    	COM126	COM126_02	전문서비스업
//	                    	COM126	COM126_03	건축기술, 엔지니어링 및 기타 과학기술 서비스업
//	                    	COM127	COM127_02	사업지원 서비스업
//	                    	COM129	COM129_01	교육 서비스업
//	                    	COM130	COM130_01	보건업
//	                    	COM130	COM130_02	사회복지 서비스업
//	                    	COM131	COM131_01	창작, 예술 및 여가관련 서비스업
//	                    	COM132	COM132_01	협회 및 단체
//	                    	COM133	COM133_01	가구내 고용활동
//	                    	COM133	COM133_02	달리 분류되지 않은 자가소비를 위한 가구의 재화 및 서비스 생산활동
	                    	 if(ocCode!=null&&!ocCode.isEmpty()&&!rowValues.get(33).isEmpty()&&i==33) {
                 	        	ComDefaultCodeVO defaultVO = new ComDefaultCodeVO();
                 	        	defaultVO.setCodeId(rowValues.get(32));
                 	        	List<CmmnDetailCode> codeDetailList = cmmUseService.selectCmmCodeDetail(defaultVO);
                 	        	boolean codeFound = false;
                                 for (Object item : codeDetailList) {
                                	 if (item instanceof CmmnDetailCode) { // ComDefaultCodeVO 타입인지 확인
                                		 CmmnDetailCode codeVO = (CmmnDetailCode) item; // ComDefaultCodeVO로 캐스팅
                                         if (rowValues.get(33).equals(codeVO.getCode())) { // CODE 값을 비교
                                             codeFound = true;
//                                             break;
                                         }
                                     }
                                 }
                                 if (!codeFound) { // 세부 코드 목록에 rowValues.get(8)이 없는 경우
                                     errors.add(rowNum + "행 " + (i + 2) + "열: 표준산업 중분류 코드를 확인해주세요.");
                                     rowSuccess = false;
                                 }		
                 	        }
	                    	if(rowValues.get(33).isEmpty()&&i==33) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 표준산업-중분류값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(rowValues.get(34).isEmpty()&&i==34) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 해양수산업-대분류값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	CmmnCodeVO vo2 = new CmmnCodeVO();
                			vo2.setSearchCondition("1");
                			vo2.setSearchKeyword(rowValues.get(34));
                			vo2.setFirstIndex(0);
                			vo2.setRecordCountPerPage(20);
//                			vo.setCodeId(rowValues.get(7));
                	        List<?> ocCode2 = cmmnCodeManageService.selectCmmnCodeList(vo2);
	                    	if(!rowValues.get(34).isEmpty()&&i==34) {
                    		
                    	        if(ocCode2 == null || ocCode2.isEmpty()) {
                    	        	errors.add(rowNum + "행 " + (i + 1) + "열: 해양수산업 대분류 코드를 확인해주세요.");
			                        rowSuccess = false;	                    	        	
                    	        }
                    	    
                    		}
	                        if(ocCode!=null&&!ocCode.isEmpty()&&!rowValues.get(35).isEmpty()&&i==35) {
                	        	ComDefaultCodeVO defaultVO = new ComDefaultCodeVO();
                	        	defaultVO.setCodeId(rowValues.get(34));
                	        	List<CmmnDetailCode> codeDetailList = cmmUseService.selectCmmCodeDetail(defaultVO);
                	        	boolean codeFound = false;
                                for (Object item : codeDetailList) {
                               	 if (item instanceof CmmnDetailCode) { // ComDefaultCodeVO 타입인지 확인
                               		CmmnDetailCode codeVO = (CmmnDetailCode) item; // ComDefaultCodeVO로 캐스팅
                                     if (rowValues.get(35).equals(codeVO.getCode())) { // CODE 값을 비교
                                         codeFound = true;
//                                         break;
                                     }
                                 }
                                }
                                if (!codeFound) { // 세부 코드 목록에 rowValues.get(35)이 없는 경우
                                    errors.add(rowNum + "행 " + (i + 2) + "열: 해양수산업 중분류 코드를 확인해주세요.");
                                    rowSuccess = false;
                                }		
                	        }
	                    	if(rowValues.get(35).isEmpty()&&i==35) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 해양수산업-중분류값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(rowValues.get(36).isEmpty()&&i==36) {
                    			errors.add(rowNum + "행 " + (i + 1) + "열: 기술명값이 비어있습니다.");
		                        rowSuccess = false;
                    		}
	                    	if(rowValues.get(37).isEmpty()&&i==37) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 연구자명값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(38).isEmpty()&&i==38) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 연구기관값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(39).isEmpty()&&i==39) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 적용분야값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(40).isEmpty()&&i==40) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 키워드값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(45).isEmpty()&&i==45) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 요악값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(rowValues.get(47).isEmpty()) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 기술완성도값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(!rowValues.get(47).isEmpty()&&!rowValues.get(47).equals("TRL1")&&!rowValues.get(47).equals("TRL2")
	                    			&&!rowValues.get(47).equals("TRL3")&&!rowValues.get(47).equals("TRL4")&&!rowValues.get(47).equals("TRL5")
	                    			&&!rowValues.get(47).equals("TRL6")&&!rowValues.get(47).equals("TRL7")&&!rowValues.get(47).equals("TRL8")
	                    			&&!rowValues.get(47).equals("TRL9")&&!rowValues.get(47).equals("ETC")&&i==47) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 기술완성도 코드를 확인해주세요.");
                                rowSuccess = false;
	                    	}
	                    
	                    	if(rowValues.get(48).isEmpty()&&rowValues.get(49).isEmpty()&&rowValues.get(50).isEmpty()
	                    			&&rowValues.get(51).isEmpty()&&rowValues.get(52).isEmpty()&&rowValues.get(53).isEmpty()) {
	                    		errors.add(rowNum + "행  거래유형값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(!rowValues.get(48).isEmpty()&&!rowValues.get(49).isEmpty()&&!rowValues.get(50).isEmpty()
	                    			&&!rowValues.get(51).isEmpty()&&!rowValues.get(52).isEmpty()&&!rowValues.get(53).isEmpty()
	                    			&&(!rowValues.get(48).equals("Y")||!rowValues.get(49).equals("Y")||!rowValues.get(50).equals("Y")
	                    			||!rowValues.get(51).equals("Y")||!rowValues.get(52).equals("Y")||!rowValues.get(53).equals("Y"))) {
	                    		errors.add(rowNum + "행  거래유형 코드를 확인해주세요.");
                                rowSuccess = false;	                    		
	                    	}
	                    	if(rowValues.get(54).isEmpty()&&i==54) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 기술료조건 값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(!rowValues.get(54).isEmpty()&&!rowValues.get(54).equals("FREE")&&!rowValues.get(54).equals("50")
	                    			&&!rowValues.get(54).equals("500")&&!rowValues.get(54).equals("1000")&&!rowValues.get(54).equals("2000")
	                    			&&!rowValues.get(54).equals("3000")&&!rowValues.get(54).equals("5000")&&!rowValues.get(54).equals("10000")
	                    			&&!rowValues.get(54).equals("ETC")&&i==54) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 기술료조건 코드를 확인해주세요.");
                                rowSuccess = false;
	                    	}
	                    	
	                    	if(rowValues.get(55).isEmpty()&&i==55) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 거래상태정보 값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	if(!rowValues.get(55).isEmpty()&&!rowValues.get(55).equals("01")&&!rowValues.get(55).equals("02")&&i==55) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 거래상태정보 코드를 확인해주세요.");
                                rowSuccess = false;
	                    	}
	                    	if(rowValues.get(56).isEmpty()&&i==56&&i==56) {
	                    		errors.add(rowNum + "행 " + (i + 1) + "열: 기술의 우수성 값이 비어있습니다.");
		                        rowSuccess = false;
	                    	}
	                    	
	                    	//파일체크 확인.
	                    	 //SMK파일

	              	  	 
	                    	if (!smkFile.isEmpty()&&!rowValues.get(57).isEmpty()&&i==57) {

	                    	    // rowValues.get(57) 값과 filesResult 비교
	                    	    if (!rowValues.get(57).isEmpty()) {
	                    	        boolean smkFileMatch = false;
	                    	        smkFilesResult = fileUtil.parseFileInf(smkFile, "SMK_", 0, "", "Globals.fileStorePath");
	                    	        for (FileVO fileVO : smkFilesResult) {
	                    	            if (fileVO.getOrignlFileNm().equals(rowValues.get(57))) {
	                    	                smkFileMatch = true;
	                    	                smkFilesResult = fileUtil.parseFileInf2(smkFile, "SMK_", 0, "", "Globals.fileStorePath", techSn);
//	                    	                fileVO.setAtchFileId(fileVO.getAtchFileId()+techSn);
	                    	                smkFileId = fileMngService.insertFileInfs(smkFilesResult);
//	                    	                break;
	                    	            }
	                    	        }
	                    	        if (!smkFileMatch) {
	                    	            errors.add(rowNum + "행 SMK 파일명이 엑셀 파일과 일치하지 않습니다.");
	                    	            rowSuccess = false;
	                    	        }else {
//	    	                    	    smkFileId = fileMngService.insertFileInfs(smkFilesResult);
	    	                    	    techInfoVO.setSmkFile(smkFileId);
	                    	        }
	                    	    }
	                    	}
	                    	// 대표 이미지 파일 체크 및 비교	          
	                    	if (!file2.isEmpty()) {

		                    	    for (int j = 58; j <= 60; j++) { // 61~65번째 열 (인덱스 60~64)
		                    	        if (rowValues.size() > j && !rowValues.get(j).isEmpty() &&i==j) {
		                    	            String targetFileName = rowValues.get(j);
		                    	            imgFilesResult = fileUtil.parseFileInf(file2, "IMG_", 0, "", "Globals.fileStorePath");
		                    	                boolean found = false;
		                    	                for (FileVO fileVO : imgFilesResult) {
		                    	                    if (fileVO.getOrignlFileNm().equals(targetFileName)) {
		                    	                        found = true;
//		                    	                        fileVO.setAtchFileId(fileVO.getAtchFileId()+techSn);
		                    	                        imgFilesResult = fileUtil.parseFileInf2(file2, "IMG_", 0, "", "Globals.fileStorePath",techSn);
		                    	                       mainImg = fileMngService.insertFileInfs(imgFilesResult);
		                    	                        break;
		                    	                    }
		                    	                }
		                    	             if (!found) {
		                    	                    errors.add(String.format("%d행 %d열 첨부 파일명 (%s)이 업로드된 파일과 일치하지 않습니다.", rowNum, j + 1, targetFileName));
		                    	                    rowSuccess = false; // 불일치 시 rowSuccess를 false로 설정		                    	                
		                    	            }else {
		    		                    	   
//		    		                    	    mainImg = fileMngService.insertFileInfs(imgFilesResult);
//		    		                    	    techInfoVO.setAttFile(attFile);
		                    	            }
		                    	        }
		                    	    }

	                    	}
	                    	

	                    	// 첨부 파일 체크 및 비교	                    	

	                    		if (!file3.isEmpty()) {

	                    	    for (int j = 61; j <= 65; j++) { // 61~65번째 열 (인덱스 60~64)
	                    	        if (rowValues.size() > j && !rowValues.get(j).isEmpty() &&i==j) {
	                    	            String targetFileName = rowValues.get(j);
	                    	            atchFilesResult = fileUtil.parseFileInf(file3, "ATT_", 0, "", "Globals.fileStorePath");
	                    	            if (atchFilesResult == null || atchFilesResult.isEmpty()) {
	                    	                errors.add(String.format("%d행 %d열 첨부 파일 (%s)이 업로드되지 않았지만, 필수는 아닙니다.", rowNum, j + 1, targetFileName));
	                    	            } else {
	                    	                boolean found = false;
	                    	                for (FileVO fileVO : atchFilesResult) {
	                    	                    if (fileVO.getOrignlFileNm().equals(targetFileName)) {
	                    	                        found = true;
//	                    	                        fileVO.setAtchFileId(fileVO.getAtchFileId()+techSn);
	                    	                        atchFilesResult = fileUtil.parseFileInf2(file3, "ATT_", 0, "", "Globals.fileStorePath",techSn);
	                    	                        attFile = fileMngService.insertFileInfs(atchFilesResult);
	                    	                        break;
	                    	                    }
	                    	                }
	                    	                if (!found) {
	                    	                    errors.add(String.format("%d행 %d열 첨부 파일명 (%s)이 업로드된 파일과 일치하지 않습니다.", rowNum, j + 1, targetFileName));
	                    	                    rowSuccess = false; // 불일치 시 rowSuccess를 false로 설정
	                    	                }else {	            	                    	    
//	            	                    	    attFile = fileMngService.insertFileInfs(atchFilesResult);
//	            	                    	    techInfoVO.setAttFile(attFile);
	                    	                }
	                    	            }
	                    	        }
	                    	    }

	                    	}
	                    	if(rowValues.get(66).equals("Y")&&i==66) {
	                    		if(rowValues.get(68).isEmpty()&&i==68) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 출원번호 값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(69).isEmpty()&&i==69) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 출원명 값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(70).isEmpty()&&i==70) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 출원인 값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(71).isEmpty()&&i==71) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 출원일자 값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    		if(rowValues.get(73).isEmpty()&&i==73) {
	                    			errors.add(rowNum + "행 " + (i + 1) + "열: 초록 값이 비어있습니다.");
			                        rowSuccess = false;
	                    		}
	                    	}
	                    	
	                    	
	                    	//기술정보 및 특허 체크 끝
	                    
	                		    //파일체크끝*/
	                    }
	                    if (rowSuccess) {
	                    	UserInfoVO userInfo2 = new UserInfoVO();
	                    	TechProjectVO techProjectVO2 = new TechProjectVO();
	                    	TechInfoVO techInfoVO2 = new TechInfoVO();
	                    	TechPantentVO techPantentVO2 = new TechPantentVO(); 
	                        try {
	                            // rowValues를 사용하여 VO에 값 설정
	                        	//거래자정보
	                        	
//	                        	userInfo2.setTechCd(techInfoVO.getTechCd());
	                        	userInfo2.setTechSn(techSn);
	                        	userInfo2.setVenCd(rowValues.get(0));//거래자구분
	                        	userInfo2.setVenOrgNm(rowValues.get(1));//기관명/기업명
	                        	userInfo2.setVenCateCd(rowValues.get(2));//기업구분
	                        	userInfo2.setVenCeoNm(rowValues.get(3));//이름(대표자)
	                        	userInfo2.setVenLicNo(rowValues.get(4));//사업자등록번호
	                        	userInfo2.setVenStDtm(rowValues.get(5));//기업설립일
	                        	userInfo2.setVenStaffCnt(rowValues.get(6));//상시 종업원 수
	                        	userInfo2.setBizCode1(rowValues.get(7));//기업 분류(대)
	                        	userInfo2.setBizCode2(rowValues.get(8));//기업 분류(중)
	                        	userInfo2.setVenOrgNm(rowValues.get(9));//이름(담당자)
	                        	userInfo2.setVenDept(rowValues.get(10));//부서명
	                        	userInfo2.setVenPos(rowValues.get(11));//직책
	                        	userInfo2.setVenEmail(rowValues.get(12));//이메일1
	                        	userInfo2.setVenEmail2(rowValues.get(13));//이메일2
	                        	userInfo2.setVenTel1(rowValues.get(14));//전화번호1
	                        	userInfo2.setVenTel2(rowValues.get(15));//전화번호2
	                        	userInfo2.setVenTel3(rowValues.get(16));//전화번호3
	                        	userInfo2.setVenHp1(rowValues.get(17));//휴대전화1
	                        	userInfo2.setVenHp2(rowValues.get(18));//휴대전화2
	                        	userInfo2.setVenHp3(rowValues.get(19));//휴대전화3
	                        	userInfo2.setVenZip(rowValues.get(20));//우편번호
	                        	userInfo2.setVenAdd1(rowValues.get(21));//도로명
	                        	userInfo2.setVenAdd2(rowValues.get(22));//상세주소
	                        	userInfo2.setAnswer1(rowValues.get(23));//이메일 답변
	                        	userInfo2.setAnswer2(rowValues.get(24));//휴대전화 답변
	                            userInfoList.add(userInfo2);
	                            
	                            //과제정보	                            
	                            int projectSn = projectIdGnrService.getNextIntegerId();	                            
	                            techProjectVO2.setTechSn(techSn);
	                            techProjectVO2.setProjectSn(projectSn);
	                            techProjectVO2.setRndType(rowValues.get(25));//연구자 과제유형
	                            techProjectVO2.setBizYearSt(rowValues.get(26));			//사업 시작연도
	                            techProjectVO2.setBizYearEt(rowValues.get(27));			//사업 끝연도
	                            techProjectVO2.setDept(rowValues.get(28));			//부처명
	                            techProjectVO2.setProjectNm(rowValues.get(29));			//과제명
	                            techProjectVO2.setProjectNo(rowValues.get(30));			//과제번호
	                            techProjectVO2.setAgency(rowValues.get(31));			//수행연구 기관명
	                            techProjectInfoList.add(techProjectVO2);
	                            //기술정보	                            
	                            techInfoVO2.setTechSn(techSn);
	                            techInfoVO2.setTechCd(techCd);
	                            techInfoVO2.setCateStLcd(rowValues.get(32));//표준산업-대분류
	                            techInfoVO2.setCateStMcd(rowValues.get(33));//표준산업-중분류
	                            techInfoVO2.setCateOcLcd(rowValues.get(34));//해양수산업-대분류
	                            techInfoVO2.setCateOcMcd(rowValues.get(35));//해양수산업-중분류
	                            techInfoVO2.setTechNm(rowValues.get(36));//기술명
	                            techInfoVO2.setResearcher(rowValues.get(37));//연구자명
	                            techInfoVO2.setOrgNm(rowValues.get(38));//연구기관
	                            techInfoVO2.setTechAplyArea(rowValues.get(39));//적용분야	                            
	                            String keyWord = rowValues.get(40);
	                            if(!rowValues.get(41).isEmpty()) {
	                            	keyWord = keyWord + "," + rowValues.get(41);
	                            }
	                            if(!rowValues.get(42).isEmpty()) {
	                            	keyWord = keyWord + "," + rowValues.get(42);
	                            }
	                            if(!rowValues.get(43).isEmpty()) {
	                            	keyWord = keyWord + "," + rowValues.get(43);
	                            }
	                            if(!rowValues.get(44).isEmpty()) {
	                            	keyWord = keyWord + "," + rowValues.get(44);
	                            }
	                            techInfoVO2.setTechKeyword(keyWord);//키워드
	                            techInfoVO2.setTechSbc(rowValues.get(45));//요약
	                            techInfoVO2.setTechRepSbc(rowValues.get(46));//대표청구항
	                            techInfoVO2.setTechPercectionCd(rowValues.get(47));//기술완성도(tr1~tr9, etc)
	                            techInfoVO2.setDealCateCd(rowValues.get(48));//거래유형(기술매매)
	                            techInfoVO2.setDealCateCd2(rowValues.get(49));//거래유형(통상실시권)
	                            techInfoVO2.setDealCateCd3(rowValues.get(50));//거래유형(독점적 통상실시권)
	                            techInfoVO2.setDealCateCd4(rowValues.get(51));//거래유형(전용실시권)
	                            techInfoVO2.setDealCateCd5(rowValues.get(52));//거래유형(기술협력)
	                            techInfoVO2.setDealCateCd6(rowValues.get(53));//거래유형(기타)
	                            techInfoVO2.setTechRoyCd(rowValues.get(54));//기술료조건
	                            techInfoVO2.setDealStCd(rowValues.get(55));//거래상태정보
	                            techInfoVO2.setTechExcellence(rowValues.get(56));//기술의 우수성
	                    	    techInfoVO2.setSmkFile(smkFileId);
		                    	techInfoVO2.setMainImg(mainImg);
	                    	    techInfoVO2.setAttFile(attFile);
	                            techInfoList.add(techInfoVO2);
	                            //특허
	                            int patentSn =pentnetIdGnrService.getNextIntegerId();
	                            if(rowValues.get(66).equals("Y")) {
		                            techPantentVO2.setTechSn(techSn);
		                            techPantentVO2.setPatentSn(patentSn);
		                            techPantentVO2.setPatentCd(rowValues.get(67));//특허(구분)
		                            techPantentVO2.setPatentAplyNo(rowValues.get(68));//특허(출원번호)68
		                            techPantentVO2.setPatentNm(rowValues.get(69));//특허(출원명)69
		                            techPantentVO2.setRgmnNm(rowValues.get(70));//특허(출원인)70
		                            techPantentVO2.setPatentDtm(rowValues.get(71));//특허(출원일자)71
		                            techPantentVO2.setPatentNtn(rowValues.get(72));//특허(국가)72
		                            techPantentVO2.setPatentAbstract(rowValues.get(73));//특허(초록)73
		                            techPantentVO2.setIpcNumber(rowValues.get(74));//IPC분류코드74
		                            techPantentInfoList.add(techPantentVO2);
	                            }
	                        } catch (Exception voEx) {
	                            errors.add(rowNum + "행 데이터 VO 매핑 중 오류 발생: " + voEx.getMessage());
	                            rowSuccess = false;
	                        }
	                    }
	                }
	                	// 기술정보
	                if (errors.isEmpty() && !techInfoList.isEmpty()) { //모두 통과했을때만 DB에 넣도록
	                    try {
	                        for (TechInfoVO techInfo : techInfoList) {
//	                        	insertTechInfo(techInfo);
	                        	techInfoDAO.insertTechInfo(techInfo);
	                            successCount++;
	                        }
	                    } catch (Exception dbEx) {
	                        failCount = techInfoList.size(); //전부 실패로 처리
	                        errors.add("DB 저장 중 오류가 발생했습니다: " + dbEx.getMessage());
	                    }
	                } else {
	                    failCount = techInfoList.size() - successCount;
	                }
	                //유저정보
	                if (errors.isEmpty() && !userInfoList.isEmpty()) { //모두 통과했을때만 DB에 넣도록
	                    try {
	                        for (UserInfoVO userInfo : userInfoList) {
	                        	userInfoService.insertTechUserInfo(userInfo);
	                            successCount++;
	                        }
	                    } catch (Exception dbEx) {
	                        failCount = techInfoList.size(); //전부 실패로 처리
	                        errors.add("DB 저장 중 오류가 발생했습니다: " + dbEx.getMessage());
	                    }
	                } else {
	                    failCount = techInfoList.size() - successCount;
	                }
	                //특허
	                if (errors.isEmpty() && !techProjectInfoList.isEmpty()) { //모두 통과했을때만 DB에 넣도록
	                    try {
	                        for (TechProjectVO techPorjectVO : techProjectInfoList) {
	                        	insertProjectInfo(techPorjectVO);
	                            successCount++;
	                        }
	                    } catch (Exception dbEx) {
	                        failCount = techProjectInfoList.size(); //전부 실패로 처리
	                        errors.add("DB 저장 중 오류가 발생했습니다: " + dbEx.getMessage());
	                    }
	                } else {
	                    failCount = techInfoList.size() - successCount;
	                }
	                if (errors.isEmpty() && !techPantentInfoList.isEmpty()) { //모두 통과했을때만 DB에 넣도록
	                    try {
	                        for (TechPantentVO techPatentVO : techPantentInfoList) {
	                        	insertPantentInfo(techPatentVO);
	                            successCount++;
	                        }
	                    } catch (Exception dbEx) {
	                        failCount = techProjectInfoList.size(); //전부 실패로 처리
	                        errors.add("DB 저장 중 오류가 발생했습니다: " + dbEx.getMessage());
	                    }
	                } else {
	                    failCount = techInfoList.size() - successCount;
	                }
	            } catch (Exception workbookEx) {
	                errors.add("엑셀 파일 형식이 잘못되었거나 파일 처리 중 오류가 발생했습니다: " + workbookEx.getMessage());
	            }
	        }
	    } catch (Exception e) {
	        errors.add("서버 오류가 발생했습니다: " + e.getMessage());
	    }

	    result.put("success", errors.isEmpty()); // errors 리스트가 비어있으면 성공
	    result.put("totalCount", techInfoList.size());
	    result.put("successCount", successCount);
	    result.put("failCount", failCount);
	    result.put("errors", errors); // 오류 목록 반환
	    
	    return result;
        
        
    }
	// 특정 컬럼의 대표 이미지 비교 메서드 (핵심!!!)
	private boolean compareImgFile(int rowNum, List<FileVO> fileVOs, List<String> rowValues, int columnIndex, List<String> errors) {
	    if (rowValues.size() > columnIndex && !rowValues.get(columnIndex).isEmpty()) { // 해당 컬럼의 값이 비어있지 않은 경우에만 비교
	        String targetFileName = rowValues.get(columnIndex); // 비교할 파일 이름
	        if (fileVOs == null || fileVOs.isEmpty() || !compareFileNames(fileVOs, targetFileName)) { // 파일이 null이거나 비어있거나, 파일이름이 일치하지 않는 경우
	             errors.add(String.format("%d행 %d열 대표 이미지 파일명 (%s)이 업로드된 파일과 일치하지 않거나, 파일이 업로드되지 않았습니다.", rowNum, columnIndex + 1, targetFileName));
	            return false; // 파일이 일치하지 않으면 false 반환 (rowSuccess에 영향을 줌)
	        }
	    }
	    return true; // 해당 컬럼의 값이 비어 있거나, 파일 이름이 일치하면 true 반환
	}
	
	// 특정 컬럼의 첨부 파일 비교 메서드 (핵심!!!)
	private boolean compareAttFile(int rowNum, List<FileVO> fileVOs, List<String> rowValues, int columnIndex, List<String> errors) {
	    if (rowValues.size() > columnIndex && !rowValues.get(columnIndex).isEmpty()) { // 해당 컬럼의 값이 비어있지 않은 경우에만 비교
	        String targetFileName = rowValues.get(columnIndex); // 비교할 파일 이름
	        if (fileVOs == null || fileVOs.isEmpty() || !compareFileNames(fileVOs, targetFileName)) { // 파일이 null이거나 비어있거나, 파일이름이 일치하지 않는 경우
	            errors.add(String.format("%d행 %d열 첨부 파일명 (%s)이 업로드된 파일과 일치하지 않거나, 파일이 업로드되지 않았습니다.", rowNum, columnIndex + 1, targetFileName));
	            return false; // 파일이 일치하지 않으면 false 반환 (rowSuccess에 영향을 줌)
	        }
	    }
	    return true; // 해당 컬럼의 값이 비어 있거나, 파일 이름이 일치하면 true 반환
	}

	// 파일 이름 비교 함수 (이전 코드와 동일)
	private boolean compareFileNames(List<FileVO> fileVOs, String targetFileName) {
	    if (fileVOs == null || targetFileName == null || targetFileName.isEmpty()) return false;
	    for (FileVO fileVO : fileVOs) {
	        if (fileVO.getOrignlFileNm().equals(targetFileName)) {
	            return true; // 파일이 일치하면 true 반환
	        }
	    }
	    return false; // 일치하는 파일을 찾지 못한 경우 false 반환
	}
	private String cellToString(Cell cell) {
	    if (cell == null) {
	        return "";
	    }
	    int cellType = cell.getCellType(); // int형으로 받음

	    switch (cellType) {
	        case Cell.CELL_TYPE_STRING: // Cell.CELL_TYPE_STRING 사용
	            return cell.getStringCellValue();
	        case Cell.CELL_TYPE_NUMERIC:
	            if (DateUtil.isCellDateFormatted(cell)) {
	                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
	                return dateFormat.format(cell.getDateCellValue());
	            } else {
	                return String.valueOf((long) cell.getNumericCellValue());
	            }
	        case Cell.CELL_TYPE_BOOLEAN:
	            return String.valueOf(cell.getBooleanCellValue());
	        case Cell.CELL_TYPE_FORMULA:
	             try {
	                return cell.getStringCellValue(); // 수식 결과가 문자열인 경우
	            } catch (IllegalStateException e) {
	                try {
	                    return String.valueOf((long) cell.getNumericCellValue()); // 수식 결과가 숫자인 경우
	                } catch (IllegalStateException e2) {
	                    return cell.getCellFormula();//수식 그대로 반환
	                }
	            }
	        case Cell.CELL_TYPE_BLANK:
	            return "";
	        default:
	            return "";
	    }
	}
	
	@Override
	public List<?> selectProjectInfo(TechProjectVO TechProjectVO, ModelMap model) throws Exception {
		
		Map<String, Object> resultMapTotal = new HashMap<String, Object>();
		List<TechProjectVO> resultList = new ArrayList<TechProjectVO>();
		
		String indexName = "project_info";
		
		ObjectMapper objectMapper = new ObjectMapper();
		SearchCommonVO searchCommonVO = new SearchCommonVO();
		String searchParam = TechProjectVO.getSearchParam();
		
		if(searchParam == null || searchParam.equals("")) {
			searchParam = "{\"indices\" : [\"" + indexName + "\"], \"fields\" : [\"PROJECT_NM\"], \"pagination\": {\"page\": 1,\"size\": 5,\"paginationBarSize\": 10}}";
		}
		
		searchCommonVO = objectMapper.readValue(searchParam.replaceAll("&quot;", "\""), SearchCommonVO.class);
		
		resultMapTotal = searchCommon.searchList(searchCommonVO, indexName);
		
		int totalCnt = (int) resultMapTotal.get("totalCnt");
		model.addAttribute("totalCnt",totalCnt);
		
		if(totalCnt > 0) {
			List<Map<String, Object>> resultMapList = (List<Map<String, Object>>) resultMapTotal.get("resultList");
			
			// 결과 출력
	        for (Map<String, Object> resultMap2 : resultMapList) {
	        	TechProjectVO vo = new TechProjectVO();
	        	for (Map.Entry<String, Object> entry : resultMap2.entrySet()) {
	        		if(entry.getKey().equals("PROJECT_NM")) {
	        			vo.setProjectNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("PROJECT_ESNTL_NO")) {
	        			vo.setProjectEsntlNo((String) entry.getValue());
	        		}else if(entry.getKey().equals("PROJECT_EXC_YEAR")) {
	        			vo.setProjectExcYear((String) entry.getValue());
	        		}else if(entry.getKey().equals("PROJECT_EXC_ORG_NM")) {
	        			vo.setProjectExcOrgNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("PROJECT_MIRYFC_NM")) {
	        			vo.setProjectMiryfcNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("GVRN_RSRCCT")) {
	        			vo.setGvrnRsrcct((String) entry.getValue());
	        		}
	        	}
	        	
	        	resultList.add(vo);
	        }
		}
		
		return resultList;
	}
	
	@Override
	public HashSet<String> selectProjectNmAutoComplete(TechProjectVO TechProjectVO) throws Exception {
		
		Map<String, Object> resultMapTotal = new HashMap<String, Object>();
		HashSet<String> resultList = new HashSet<String>();
		
		String indexName = "project_info";
		
		ObjectMapper objectMapper = new ObjectMapper();
		String searchParam = TechProjectVO.getSearchParam();
		
		SearchCommonVO searchAutoCompleteVO = new SearchCommonVO();
		
		searchAutoCompleteVO = objectMapper.readValue(searchParam.replaceAll("&quot;", "\""), SearchCommonVO.class);
		
		resultMapTotal = searchCommon.searchList(searchAutoCompleteVO, indexName);
		
		int totalCnt = (int) resultMapTotal.get("totalCnt");
		
		if(totalCnt > 0) {
			List<Map<String, Object>> resultMapList = (List<Map<String, Object>>) resultMapTotal.get("resultList");
			
			// 결과 출력
	        for (Map<String, Object> resultMap2 : resultMapList) {
	        	String projectNm = "";
	        	for (Map.Entry<String, Object> entry : resultMap2.entrySet()) {
	        		if(entry.getKey().equals("PROJECT_NM")) {
	        			projectNm = (String) entry.getValue();
	        		}
	        	}
	        	resultList.add(projectNm);
	        }
		}
		
		return resultList;
	}

	@Override
	public void updateSmkText(TechInfoVO techInfoVO) throws Exception {
		techInfoDAO.updateSmkText(techInfoVO);
	}

	@Override
	public List<?> selectTechInfoListSearch(TechInfoVO techInfoVO, ModelMap model) throws Exception {
		
		Map<String, Object> resultMapTotal = new HashMap<String, Object>();
		List<TechInfoVO> resultList = new ArrayList<TechInfoVO>();
		
		String indexName = "tech_info";
		
		ObjectMapper objectMapper = new ObjectMapper();
		SearchCommonVO searchCommonVO = new SearchCommonVO();
		String searchParam = techInfoVO.getSearchParam();
		
		if(searchParam == null || searchParam.equals("")) {
			if(techInfoVO.getTechCd().equals("sales")) {
				//개발
				searchParam = "{\"indices\" : [\"" + indexName + "\"], \"fields\" : [\"TECH_NM\"],  \"query\" : {\"text\" : \"\",\"includeText\":\"\",\"excludeText\":\"\",\"exactText\":[],\"operator\":\"\"}, \"filter\" : [{\"filterType\" : \"TERMS\",\"field\" : \"TECH_CD\",\"text\" : [\"@REPLACE_TECH_CD@\"],\"gte\" : null,\"lte\" : null,\"format\" : null}], \"pagination\": {\"page\": 1,\"size\": 5,\"paginationBarSize\": 10}, \"sort\": [{\"field\": \"DEAL_ST_CD.keyword\", \"order\": \"ASC\", \"indices\": [\"" + indexName + "\"]}, {\"field\": \"TECH_SN\", \"order\": \"DESC\", \"indices\": [\"" + indexName + "\"]}]}";				
				//운영
				//searchParam = "{\"indices\" : [\"" + indexName + "\"], \"fields\" : [\"TECH_NM\"],   \"filter\" : [{\"filterType\" : \"TERMS\",\"field\" : \"TECH_CD\",\"text\" : [\"@REPLACE_TECH_CD@\"],\"gte\" : null,\"lte\" : null,\"format\" : null}], \"pagination\": {\"page\": 1,\"size\": 5,\"paginationBarSize\": 10}, \"sort\": [{\"field\": \"DEAL_ST_CD.keyword\", \"order\": \"ASC\", \"indices\": [\"" + indexName + "\"]}, {\"field\": \"TECH_SN\", \"order\": \"DESC\", \"indices\": [\"" + indexName + "\"]}]}";				
			
			}else {
				//개발
				searchParam = "{\"indices\" : [\"" + indexName + "\"], \"fields\" : [\"TECH_NM\"],  \"query\" : {\"text\" : \"\",\"includeText\":\"\",\"excludeText\":\"\",\"exactText\":[],\"operator\":\"\"}, \"filter\" : [{\"filterType\" : \"TERMS\",\"field\" : \"TECH_CD\",\"text\" : [\"@REPLACE_TECH_CD@\"],\"gte\" : null,\"lte\" : null,\"format\" : null}], \"pagination\": {\"page\": 1,\"size\": 5,\"paginationBarSize\": 10}, \"sort\": [{\"field\": \"TECH_SN\", \"order\": \"DESC\", \"indices\": [\"" + indexName + "\"]}]}";		
				//운영
				//searchParam = "{\"indices\" : [\"" + indexName + "\"], \"fields\" : [\"TECH_NM\"],  \"filter\" : [{\"filterType\" : \"TERMS\",\"field\" : \"TECH_CD\",\"text\" : [\"@REPLACE_TECH_CD@\"],\"gte\" : null,\"lte\" : null,\"format\" : null}], \"pagination\": {\"page\": 1,\"size\": 5,\"paginationBarSize\": 10}, \"sort\": [{\"field\": \"TECH_SN\", \"order\": \"DESC\", \"indices\": [\"" + indexName + "\"]}]}";		
			}
			

		
		}
		
		searchParam = searchParam.replaceAll("@REPLACE_TECH_CD@", techInfoVO.getTechCd());
		
		searchCommonVO = objectMapper.readValue(searchParam.replaceAll("&quot;", "\""), SearchCommonVO.class);
		
		resultMapTotal = searchCommon.searchList(searchCommonVO, indexName);
		
		int totalCnt = (int) resultMapTotal.get("totalCnt");
		model.addAttribute("totalCnt",totalCnt);

		if(totalCnt > 0) {
			List<Map<String, Object>> resultMapList = (List<Map<String, Object>>) resultMapTotal.get("resultList");
			
			// 결과 출력
	        for (Map<String, Object> resultMap2 : resultMapList) {
	        	TechInfoVO vo = new TechInfoVO();
	        	for (Map.Entry<String, Object> entry : resultMap2.entrySet()) {
	        		if(entry.getKey().equals("CATESTLNM")) {
	        			vo.setCateStLNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_SBC")) {
	        			vo.setTechSbc((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_APLY_AREA")) {
	        			vo.setTechAplyArea((String) entry.getValue());
	        		}else if(entry.getKey().equals("CATE_ST_MCD")) {
	        			vo.setCateStMcd((String) entry.getValue());
	        		}else if(entry.getKey().equals("PATENT_DTM")) {
	        			vo.setPatentDtm((String) entry.getValue());
	        		}else if(entry.getKey().equals("CATESTLNM")) {
	        			vo.setCateStLNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("CATESTMNM")) {
	        			vo.setCateStMNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("CATEOCLNM")) {
	        			vo.setCateOcLNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("CATEOCMNM")) {
	        			vo.setCateOcMNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_NM")) {
	        			vo.setVenNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_CD")) {
	        			vo.setVenCd((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_ORG_NM")) {
	        			vo.setVenOrgNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_CATE_CD")) {
	        			vo.setVenCateCd((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_LIC_NO")) {
	        			vo.setVenLicNo((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_ZIP")) {
	        			vo.setVenZip((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_ADD")) {
	        			vo.setVenAdd((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_ST_DT")) {
	        			vo.setVenStDtm((String) entry.getValue());
	        		}else if(entry.getKey().equals("VEN_STAFF_CNT")) {
	        			vo.setVenStaffCnt((String) entry.getValue());
	        		}else if(entry.getKey().equals("DEAL_CATE_CD")) {
	        			vo.setDealCateCd((String) entry.getValue());
	        		}else if(entry.getKey().equals("DEAL_CATE_CD2")) {
	        			vo.setDealCateCd2((String) entry.getValue());
	        		}else if(entry.getKey().equals("DEAL_CATE_CD3")) {
	        			vo.setDealCateCd3((String) entry.getValue());
	        		}else if(entry.getKey().equals("DEAL_CATE_CD4")) {
	        			vo.setDealCateCd4((String) entry.getValue());
	        		}else if(entry.getKey().equals("DEAL_CATE_CD5")) {
	        			vo.setDealCateCd5((String) entry.getValue());
	        		}else if(entry.getKey().equals("DEAL_CATE_CD6")) {
	        			vo.setDealCateCd6((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_CD_NM")) {
	        			vo.setTechCdNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_NM")) {
	        			vo.setTechNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_EXCELLENCE")) {
	        			vo.setTechExcellence((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_APLY_DTM")) {
	        			vo.setTechAplyDtm((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_REP_SBC")) {
	        			vo.setTechRepSbc((String) entry.getValue());
	        		}else if(entry.getKey().equals("ORG_NM")) {
	        			vo.setOrgNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("CATE_OC_MCD")) {
	        			vo.setCateOcMcd((String) entry.getValue());
	        		}else if(entry.getKey().equals("CATE_OC_LCD")) {
	        			vo.setCateOcLcd((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_CD")) {
	        			vo.setTechCd((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_SN")) {
	        			vo.setTechSn(Integer.parseInt((String) entry.getValue()));
	        		}else if(entry.getKey().equals("CATE_ST_LCD")) {
	        			vo.setCateStLcd((String) entry.getValue());
	        		}else if(entry.getKey().equals("TECH_PERCECTION_NM")) {
	        			vo.setTechPercectionNm((String) entry.getValue());
	        		}else if(entry.getKey().equals("SMK_FILE")) {
	        			vo.setSmkFile((String) entry.getValue());
	        		}else if(entry.getKey().equals("DEAL_ST_CD")) {
	        			vo.setDealStCd((String) entry.getValue());
	        		}else if(entry.getKey().equals("REG_DTM")) {
	        			vo.setRegDtm((String) entry.getValue());
	        		}else if(entry.getKey().equals("_file")) {
	        			List<Map<String, Object>> fileInfo = (List<Map<String, Object>>) entry.getValue();
	        			String value = "";
	        			if(fileInfo != null) {
	        				for (Map<String, Object> fileItem : fileInfo) {
	        					for (Map.Entry<String, Object> fileEntry : fileItem.entrySet()) {
	        						if(entry.getKey().equals("_content")) {
	        							value = (String) entry.getValue();
	        						}
	        					}
	        				}
	    				}
	        		}
	        	}
	        	vo.setMberId(techInfoVO.getMberId());
	        	int myScrapCnt = 0 ;
	        	if(vo.getMberId() != null) {
	        		myScrapCnt = techInfoDAO.selectTechInfoMyScrapCnt(vo); 
	        	}
	        	vo.setMyScrapCnt(myScrapCnt);
	        	resultList.add(vo);
	        }
		}
		
		return resultList;
	}
}
