Attribute VB_Name = "SetTokArray"
Option Explicit

'-------------------------------
' Constants for Token processing
'-------------------------------
'Public Const TK_none = 0
'Public Const TK_uint = 1
'Public Const TK_str = 2
'Public Const TK_dword = 3
'Public Const TK_float = 4
'Public Const TK_uint4float = 5
'Public Const TK_2sint3float = 6
'Public Const TK_sint = 7
'Public Const TK_2uint2float = 8
'Public Const TK_uintfloat = 9
'Public Const TK_dworduint = 10
'Public Const TK_tokuintfloat = 11
'Public Const TK_uintuint = 12
'Public Const TK_uintfloatdword = 13
'Public Const TK_dworduintfloat = 14
'Public Const TK_uintfloat6 = 15
'Public Const TK_mixed1 = 16
'Public Const TK_uintplus = 17
'Public Const TK_uintplusfloat = 18
'Public Const TK_mixed3 = 19
'Public Const TK_mixed4 = 20
'Public Const TK_uintnocr = 21
'Public Const TK_mixed2 = 22
'Public Const TK_tokfloat = 23
'Public Const TK_struint = 24
'Public Const TK_dwordfloat = 25
'Public Const TK_buffer = 26
'Public Const TK_somefloat = 27
'
'Public Const TK_literal = 0
'Public Const TK_level = 1
'Public Const TK_embedded = 2
'Public Const TK_type = 3
'Public Const TK_count = 4
'Public Const TK_precis = 5
'
'Public Const TK_embedded_yes = 1
'Public Const TK_embedded_no = 0

'Public Const S_Type = 0
'Public Const W_Type = 1
'Public Const T_Type = 2

Public Tokenz As New colTok

Public Tokens As New colTok

Public Sub Init_Tokenz()

    Call Initz_S
    Call Initz_W
    Call Initz_T
    
End Sub

Private Sub Initz_S()

    Call Set_Entry(S_Type, 2, "point", 2, 0, TK_float, 3, 6)
    Call Set_Entry(S_Type, 3, "vector", 3, 0, TK_float, 3, 6)
    Call Set_Entry(S_Type, 5, "normals", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 6, "normal_idxs", 9, 0, TK_uintuint, 2, 1)
    Call Set_Entry(S_Type, 7, "points", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 8, "uv_point", 2, 0, TK_float, 2, 6)
    Call Set_Entry(S_Type, 9, "uv_points", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 10, "colour", 2, 0, TK_float, 4, 3)
    Call Set_Entry(S_Type, 11, "colours", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 13, "image", 2, 0, TK_str, 1, 0)
    Call Set_Entry(S_Type, 14, "images", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 15, "texture", 2, 0, TK_uintfloatdword, 2, 1)
    Call Set_Entry(S_Type, 16, "textures", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 17, "light_material", 2, 0, TK_dworduintfloat, 1, 4)
    Call Set_Entry(S_Type, 18, "light_materials", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 19, "linear_key", 8, 0, TK_uintfloat6, 1, 3)
    Call Set_Entry(S_Type, 20, "tcb_key", 8, 0, TK_uintfloat6, 1, 9)
    Call Set_Entry(S_Type, 21, "linear_pos", 7, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 22, "tcb_pos", 7, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 23, "slerp_rot", 8, 0, TK_uintfloat6, 1, 4)
    Call Set_Entry(S_Type, 24, "tcb_rot", 7, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 25, "controllers", 6, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 26, "anim_node", 5, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 27, "anim_nodes", 4, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 28, "animation", 3, TK_embedded_yes, TK_uint, 2, 0)
    Call Set_Entry(S_Type, 29, "animations", 2, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 31, "lod_controls", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 32, "lod_control", 2, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 33, "distance_levels_header", 3, 0, TK_uintplusfloat, 1, 6)
    Call Set_Entry(S_Type, 34, "distance_level_header", 5, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 35, "dlevel_selection", 6, 0, TK_float, 1, 6)
    Call Set_Entry(S_Type, 36, "distance_levels", 3, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 37, "distance_level", 4, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 38, "sub_objects", 5, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 39, "sub_object", 6, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 40, "sub_object_header", 7, 0, TK_mixed3, 0, 0)
    Call Set_Entry(S_Type, 41, "geometry_info", 8, TK_embedded_yes, TK_uint, 10, 0)
    Call Set_Entry(S_Type, 42, "geometry_nodes", 9, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 43, "geometry_node", 10, TK_embedded_yes, TK_uint, 5, 0)
    Call Set_Entry(S_Type, 44, "geometry_node_map", 9, 0, TK_uintuint, 1, 7)
    Call Set_Entry(S_Type, 45, "cullable_prims", 11, 0, TK_uint, 3, 0)
    Call Set_Entry(S_Type, 46, "vtx_state", 2, 0, TK_mixed1, 0, 1)
    Call Set_Entry(S_Type, 47, "vtx_states", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 48, "vertex", 8, TK_embedded_yes, TK_mixed4, 1, 1)
    Call Set_Entry(S_Type, 49, "vertex_uvs", 9, 0, TK_uintuint, 1, 1)
    Call Set_Entry(S_Type, 50, "vertices", 7, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 51, "vertex_set", 8, 0, TK_uint, 3, 0)
    Call Set_Entry(S_Type, 52, "vertex_sets", 7, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 53, "primitives", 7, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 54, "prim_state", 2, 0, TK_mixed2, 1, 1)
    Call Set_Entry(S_Type, 55, "prim_states", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 56, "prim_state_idx", 8, 0, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 60, "indexed_trilist", 8, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 61, "tex_idxs", 3, 0, TK_uintplus, 1, 1)
    Call Set_Entry(S_Type, 63, "vertex_idxs", 9, 0, TK_uintuint, 1, 1)
    Call Set_Entry(S_Type, 64, "flags", 9, 0, TK_uintuint, 1, 3)
    Call Set_Entry(S_Type, 65, "matrix", 2, 0, TK_float, 12, 3)
    Call Set_Entry(S_Type, 66, "matrices", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 67, "hierarchy", 6, 0, TK_uintuint, 1, 7)
    Call Set_Entry(S_Type, 68, "volumes", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 69, "vol_sphere", 2, TK_embedded_yes, TK_tokfloat, 1, 6)
    Call Set_Entry(S_Type, 70, "shape_header", 1, 0, TK_dword, 2, 1)
    Call Set_Entry(S_Type, 71, "shape", 0, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 72, "shader_names", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 74, "texture_filter_names", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 76, "sort_vectors", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 79, "light_model_cfgs", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 80, "light_model_cfg", 2, TK_embedded_yes, TK_dword, 1, 0)
    Call Set_Entry(S_Type, 81, "uv_ops", 3, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 84, "uv_op_copy", 4, 0, TK_uint, 2, 0)
    Call Set_Entry(S_Type, 104, "subobject_shaders", 8, 0, TK_uintuint, 1, 1)
    Call Set_Entry(S_Type, 105, "subobject_light_cfgs", 8, 0, TK_uintuint, 1, 1)
    Call Set_Entry(S_Type, 125, "named_filter_mode", 2, 0, TK_str, 1, 0)
    Call Set_Entry(S_Type, 129, "named_shader", 2, 0, TK_str, 1, 0)
    Call Set_Entry(S_Type, 208, "uv_op_embossbump", 4, 0, TK_uintfloat6, 2, 1)
    Call Set_Entry(S_Type, 209, "uv_op_nonuniformescale", 4, 0, TK_uintfloat6, 2, 2)
    Call Set_Entry(S_Type, 210, "uv_op_reflectmap", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 211, "uv_op_reflectmapfull", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 212, "uv_op_share", 4, 0, TK_uint, 2, 0)
    Call Set_Entry(S_Type, 213, "uv_op_specularmap", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 214, "uv_op_spheremap", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 215, "uv_op_spheretmapfull", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 216, "uv_op_transform", 4, 0, TK_uintfloat6, 2, 6)
    Call Set_Entry(S_Type, 217, "uv_op_uniformscale", 4, 0, TK_uintfloat6, 2, 1)
    Call Set_Entry(S_Type, 218, "uv_op_user_nonuniformscale", 4, 0, TK_uint, 3, 0)
    Call Set_Entry(S_Type, 219, "uv_op_user_transform", 4, 0, TK_uint, 2, 0)
    Call Set_Entry(S_Type, 220, "uv_op_user_uniformscale", 4, 0, TK_uint, 3, 0)
    Call Set_Entry(S_Type, 221, "uv_opcopy", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(S_Type, 222, "indexed_line_list", 8, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(S_Type, 223, "point_list", 8, 0, TK_uint, 2, 0)
    
End Sub

Private Sub Initz_T()
    
    Call Set_Entry(T_Type, 136, "terrain", 0, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(T_Type, 137, "terrain_errthreshold_scale", 1, 0, TK_float, 1, 6)
    Call Set_Entry(T_Type, 138, "terrain_alwaysselect_maxdist", 1, 0, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 139, "terrain_samples", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(T_Type, 140, "terrain_nsamples", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 141, "terrain_sample_rotation", 2, 0, TK_float, 1, 6)
    Call Set_Entry(T_Type, 142, "terrain_sample_floor", 2, 0, TK_float, 1, 3)
    Call Set_Entry(T_Type, 143, "terrain_sample_scale", 2, 0, TK_float, 1, 3)
    Call Set_Entry(T_Type, 144, "terrain_sample_size", 2, 0, TK_float, 1, 3)
    Call Set_Entry(T_Type, 145, "terrain_sample_fbuffer", 2, 0, TK_str, 1, 0)
    Call Set_Entry(T_Type, 146, "terrain_sample_ybuffer", 2, 0, TK_str, 1, 0)
    Call Set_Entry(T_Type, 147, "terrain_sample_ebuffer", 2, 0, TK_str, 1, 0)
    Call Set_Entry(T_Type, 148, "terrain_sample_nbuffer", 2, 0, TK_str, 1, 0)
    Call Set_Entry(T_Type, 149, "terrain_sample_cbuffer", 2, 0, TK_str, 1, 0)
    Call Set_Entry(T_Type, 150, "terrain_sample_dbuffer", 2, 0, TK_str, 1, 0)
    Call Set_Entry(T_Type, 151, "terrain_shaders", 1, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 152, "terrain_shader", 2, TK_embedded_yes, TK_str, 1, 0)
    Call Set_Entry(T_Type, 153, "terrain_texslots", 3, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 154, "terrain_texslot", 4, 0, TK_struint, 1, 2)
    Call Set_Entry(T_Type, 155, "terrain_uvcalcs", 3, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 156, "terrain_uvcalc", 4, 0, TK_uintfloat, 3, 1)
    Call Set_Entry(T_Type, 157, "terrain_patches", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(T_Type, 158, "terrain_patchsets", 2, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 159, "terrain_patchset", 3, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(T_Type, 160, "terrain_patchset_distance", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 161, "terrain_patchset_npatches", 4, 0, TK_uint, 1, 0)
    Call Set_Entry(T_Type, 163, "terrain_patchset_patches", 4, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(T_Type, 164, "terrain_patchset_patch", 5, 0, TK_dwordfloat, 15, 6)
    Call Set_Entry(T_Type, 251, "terrain_water_height_offset", TK_embedded_yes, 0, TK_somefloat, 4, 3)
    Call Set_Entry(T_Type, 281, "terrain_sample_asbuffer", 2, 0, TK_buffer, 0, 0)
    Call Set_Entry(T_Type, 282, "terrain_sample_usbuffer", 2, 0, TK_buffer, 0, 0)

End Sub

Private Sub Initz_W()
    
    Call Set_Entry(W_Type, 3, "Static", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 5, "TrackObj", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 6, "Dyntrack", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 8, "Forest", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 9, "Telepole", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 11, "CollideObject", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 17, "Signal", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 56, "Gantry", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 57, "CarSpawner", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 59, "Pickup", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 60, "Platform", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 61, "Siding", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 62, "LevelCr", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 63, "Transfer", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 64, "Speedpost", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 65, "Hazard", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 75, "Tr_Worldfile", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 76, "Tr_Watermark", 1, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 95, "FileName", 2, 0, TK_str, 1, 0)
    Call Set_Entry(W_Type, 97, "Position", 2, 0, TK_float, 3, 3)
    Call Set_Entry(W_Type, 98, "Direction", 2, 0, TK_float, 3, 6)
    Call Set_Entry(W_Type, 99, "MaxVisDistance", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 101, "StaticDetailLevel", 2, 0, TK_sint, 1, 0)
    Call Set_Entry(W_Type, 104, "StaticFlags", 2, 0, TK_dword, 1, 0)
    Call Set_Entry(W_Type, 105, "CollideFlags", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 106, "CollideFunction", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 108, "UiD", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 109, "TrackSections", 2, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 110, "TrackSection", 3, 0, TK_tokuintfloat, 1, 2)
    Call Set_Entry(W_Type, 119, "SectionIdx", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 120, "SectionCurve", 4, 0, TK_uintnocr, 1, 0)
    Call Set_Entry(W_Type, 124, "JNodePosn", 2, 0, TK_2sint3float, 2, 3)
    Call Set_Entry(W_Type, 158, "SignalSubObj", 2, 0, TK_dword, 1, 0)
    Call Set_Entry(W_Type, 186, "SignalUnits", 2, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 187, "SignalUnit", 3, TK_embedded_yes, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 193, "Elevation", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 200, "Population", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 201, "Area", 2, 0, TK_float, 2, 3)
    Call Set_Entry(W_Type, 203, "ScaleRange", 2, 0, TK_float, 2, 3)
    Call Set_Entry(W_Type, 204, "StartPosition", 2, 0, TK_float, 3, 3)
    Call Set_Entry(W_Type, 205, "EndPosition", 2, 0, TK_float, 3, 3)
    Call Set_Entry(W_Type, 206, "StartType", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 207, "EndType", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 277, "StartDirection", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 278, "EndDirection", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 279, "ViewDbSphere", 1, TK_embedded_yes, TK_none, 0, 0)
    Call Set_Entry(W_Type, 280, "Radius", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 281, "NoDirLight", 2, 0, TK_uint, 1, 1)
    Call Set_Entry(W_Type, 283, "VDbId", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 284, "VDbIdCount", 1, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 298, "Matrix3x3", 2, 0, TK_float, 9, 6)
    Call Set_Entry(W_Type, 622, "TRItemId", 2, 0, TK_uint, 2, 0)
    Call Set_Entry(W_Type, 645, "QDirection", 2, 0, TK_float, 4, 6)
    Call Set_Entry(W_Type, 791, "Config", 2, 0, TK_uint, 1, 0)
    Call Set_Entry(W_Type, 807, "PlatformData", 2, 0, TK_dword, 1, 0)
    Call Set_Entry(W_Type, 811, "SpeedRange", 2, 0, TK_float, 2, 3)
    Call Set_Entry(W_Type, 812, "PickupType", 2, 0, TK_uint, 2, 0)
    Call Set_Entry(W_Type, 813, "PickupAnimData", 2, 0, TK_uintfloat, 1, 1)
    Call Set_Entry(W_Type, 814, "PickupCapacity", 2, 0, TK_float, 2, 3)
    Call Set_Entry(W_Type, 816, "CarFrequency", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 817, "CarAvSpeed", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 820, "SidingData", 2, 0, TK_dword, 1, 0)
    Call Set_Entry(W_Type, 822, "LevelCrParameters", 2, 0, TK_float, 2, 3)
    Call Set_Entry(W_Type, 823, "LevelCrData", 2, 0, TK_dworduint, 1, 1)
    Call Set_Entry(W_Type, 824, "LevelCrTiming", 2, 0, TK_float, 3, 3)
    Call Set_Entry(W_Type, 831, "Speed_Sign_Shape", 2, 0, TK_uint4float, 1, 4)
    Call Set_Entry(W_Type, 834, "Speed_Digit_Tex", 2, 0, TK_str, 1, 0)
    Call Set_Entry(W_Type, 839, "Speed_Text_Size", 2, 0, TK_float, 3, 2)
    Call Set_Entry(W_Type, 852, "Width", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 853, "Height", 2, 0, TK_float, 1, 3)
    Call Set_Entry(W_Type, 854, "TreeTexture", 2, 0, TK_str, 1, 0)
    Call Set_Entry(W_Type, 855, "TreeSize", 2, 0, TK_float, 2, 3)
    Call Set_Entry(W_Type, 1231, "CrashProbability", 2, 0, TK_float, 1, 3)
    
End Sub

Private Sub Set_Entry(kind As Long, p1 As Long, p2 As String, p3 As Long, p4 As Long, p5 As Long, p6 As Long, p7 As Long)
    
    Tokenz.Add LCase(p2), p1, p5, CBool(p4), p6, p7, LCase(p2)

End Sub

