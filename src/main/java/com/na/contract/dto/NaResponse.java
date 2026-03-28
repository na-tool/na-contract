package com.na.contract.dto;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author pg
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class NaResponse<T> {
    private Integer code;
    private String msg;
    private String tradeNo;
    private T data;
}
