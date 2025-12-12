from PIL import Image
from fairypptx import object_utils


def replace(src, dst, **kwargs): 
    """Replace `dst` 

    Args:
        src: Shape.
        dst: which replace `src`. 
        **kwargs:  keyword arguments of `Shape.make(dst)` (when `dst` is not Shape) 


    Policy - What  `replace` means? -
    ---------------------------------
    It is important to `dst` can behave like `src`, after this function.
    (Since normally you ask your friend to replace some machines, 
     then you expect them to work as before? )

    Maybe what should be transferred depends on type of `dst`.

    * `AnimationEffect` (Not Yet implemented)
    * 
    
    """
    # For importing hierarchy, import is put here.
    from fairypptx import Shape
    """Replace `src` with `dst`. 
    """
    if not isinstance(src, Shape):
        raise TypeError(f"Type of the first must be Shape, but {type(src)}")

    if not isinstance(dst, Shape):
        dst_shape = Shape.make(dst, **kwargs)
    else:
        dst_shape = dst
    assert isinstance(dst_shape, Shape)
    
    props = ["Top", "Left", "Width", "Height"]

    # Handling of optional attributes for 
    if isinstance(dst, Image.Image):
        props += ["LockAspectRatio"]
        props += ["line"] # Line's information.
    else:
        pass
    stored = {prop: object_utils.getattr(src, prop) for prop in props}
    for key, value in stored.items():
        object_utils.setattr(dst_shape, key, value)

    src.delete()
    return dst_shape
    

if __name__ == "__main__":
    pass
