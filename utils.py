#!/usr/bin/env python3
"""
Utility functions for the Document Proofreader
"""

def calculate_costs(total_input_tokens, output_tokens, cached_percentage=0):
    """
    Calculate API costs based on current pricing
    
    Parameters:
    -----------
    total_input_tokens : int
        Total number of input tokens
    output_tokens : int
        Total number of output tokens
    cached_percentage : float
        Percentage of input tokens that were cached (0-100)
        
    Returns:
    --------
    dict
        Dictionary with cost breakdown
    """
    from config import Config
    
    # Calculate non-cached vs cached tokens
    cached_tokens = int(total_input_tokens * (cached_percentage / 100))
    new_tokens = total_input_tokens - cached_tokens
    
    # Calculate costs
    new_input_cost = (new_tokens / 1000000) * Config.COST_TEXT_INPUT
    cached_input_cost = (cached_tokens / 1000000) * Config.COST_CACHED_INPUT
    output_cost = (output_tokens / 1000000) * Config.COST_OUTPUT
    total_cost = new_input_cost + cached_input_cost + output_cost
    
    return {
        "input_tokens": total_input_tokens,
        "output_tokens": output_tokens,
        "cached_tokens": cached_tokens,
        "new_tokens": new_tokens,
        "cached_percentage": cached_percentage,
        "new_input_cost": new_input_cost,
        "cached_input_cost": cached_input_cost,
        "output_cost": output_cost,
        "total_cost": total_cost
    }