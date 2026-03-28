package com.na.contract.dto;

import com.na.common.utils.NaCommonUtil;
import com.na.common.utils.NaIDUtil;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.lang3.StringUtils;

/**
 * @author pg
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NaWordToPdfDTO {
    private String base64; // word
    private String absPath; // pdf 临时存储路径
    private String libreoffice; // libreoffice 可执行文件

    public String getAbsPath() {
        if(StringUtils.isEmpty(absPath)){
            if(NaCommonUtil.isWindows()){
                return "C:\\na\\contract\\temp\\" + NaIDUtil.ID(7) + ".pdf";
            }else {
                return "/na/contract/temp/" + NaIDUtil.ID(7) + ".pdf";
            }
        }
        return absPath;
    }

    public String getLibreoffice() {
        if(StringUtils.isEmpty(libreoffice)){
            if(NaCommonUtil.isWindows()){
                return "C:\\Program Files\\LibreOffice\\program\\soffice.exe";
            }else {
                return "/opt/libreoffice25.8/program/soffice";
            }
        }
        return libreoffice;
    }

    public boolean isPdfFile() {
        return getAbsPath() != null && getAbsPath().toLowerCase().endsWith(".pdf");
    }
}
