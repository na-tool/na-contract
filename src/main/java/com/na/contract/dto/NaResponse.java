package com.na.contract.dto;

import lombok.Data;

/**
 * @author pg
 */
@Data
public class NaResponse<T> {
    private Integer code;
    private String msg;
    private String tradeNo;
    private T data;
}
