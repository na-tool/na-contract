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
public class NaWordToPdfDTO {
    private String base64;

}
